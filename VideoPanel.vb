' VideoPanel.vb
' 开源协议：MIT
' 依赖项（通过 NuGet 安装）：
'   NAudio
'   Newtonsoft.Json
'   Vortice.D3DCompiler
'   Vortice.Direct3D11
'   Vortice.DXGI
'   Vortice.Mathematics
'   System.Threading.Tasks.Extensions (可选)

Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports NAudio.Wave
Imports Newtonsoft.Json.Linq
Imports Vortice.D3DCompiler
Imports Vortice.Direct3D11
Imports Vortice.DXGI
Imports Vortice.Mathematics
Imports D3D11 = Vortice.Direct3D11
Imports DXGI = Vortice.DXGI

Public Class VideoPanel
    Inherits Panel

    ' ---------- 辅助类 ----------
    Private Class VideoFrame
        Public Property Data As Byte()
        Public Property PTS As TimeSpan
    End Class

    ' ---------- 公共枚举 ----------
    Public Enum RendererType
        Auto = 0
        GDI = 1
        Direct3D11 = 3
    End Enum

    ' ---------- 事件 ----------
    Public Event PlayStateChanged(ByVal isPlaying As Boolean)
    Public Event PositionChanged(ByVal position As TimeSpan)
    Public Event DurationChanged(ByVal duration As TimeSpan)
    Public Event MediaLoaded(ByVal info As MediaInfo)
    Public Event ErrorOccurred(ByVal message As String)
    Public Event MediaEnded()
    Public Event RendererChanged(ByVal renderer As RendererType)

    ' ---------- 媒体信息 ----------
    Public Class MediaInfo
        Public Property FilePath As String
        Public Property Duration As TimeSpan
        Public Property VideoWidth As Integer
        Public Property VideoHeight As Integer
        Public Property VideoFPS As Double
        Public Property AudioSampleRate As Integer
        Public Property AudioChannels As Integer
        Public Property Artist As String
        Public Property Album As String
    End Class

    ' ---------- 私有字段 ----------
    Private _frameQueue As New Concurrent.ConcurrentQueue(Of VideoFrame)
    Private Const MaxQueueSize As Integer = 30

    Private _ffmpegPath As String = "ffmpeg.exe"          ' 可配置
    Private _ffprobePath As String = "ffprobe.exe"

    Private _currentFile As String
    Private _isPlaying As Boolean
    Private _isPaused As Boolean
    Private _position As TimeSpan
    Private _duration As TimeSpan
    Private _volume As Single = 0.8F
    Private _playbackRate As Single = 1.0F
    Private _isMuted As Boolean

    Private _frameWidth As Integer = 640
    Private _frameHeight As Integer = 360
    Private _videoFPS As Double = 30.0
    Private _hasVideo As Boolean = True

    Private _ffmpegVideoProcess As Process
    Private _ffmpegAudioProcess As Process
    Private _cts As CancellationTokenSource
    Private _videoThread As Thread
    Private _audioThread As Thread
    Private _videoStream As Stream
    Private _audioStream As Stream

    Private _waveOut As WaveOutEvent
    Private _audioProvider As BufferedWaveProvider
    Private _audioFormat As WaveFormat

    Private _playTimer As Timer
    Private _renderTimer As Timer
    Private _playStartTime As DateTime?
    Private _basePosition As TimeSpan
    Private _pauseEvent As New ManualResetEvent(True)

    ' 音视频同步
    Private _videoClockBase As TimeSpan
    Private _decodedFrameCount As Long
    Private _syncFrameDuration As Double
    Private _avSyncOffset As Double = 0.0

    Private _renderer As RendererType = RendererType.Auto

    ' D3D11 资源
    Private ReadOnly _d3dLock As New Object()
    Private _d3dDevice As ID3D11Device
    Private _d3dContext As ID3D11DeviceContext
    Private _swapChain As IDXGISwapChain
    Private _renderTargetView As ID3D11RenderTargetView
    Private _texture2D As ID3D11Texture2D
    Private _shaderResourceView As ID3D11ShaderResourceView
    Private _vertexShader As ID3D11VertexShader
    Private _pixelShader As ID3D11PixelShader
    Private _inputLayout As ID3D11InputLayout
    Private _vertexBuffer As ID3D11Buffer
    Private _samplerState As ID3D11SamplerState
    Private _rasterState As ID3D11RasterizerState
    Private _frameDataForD3D As Byte()
    Private ReadOnly _frameDataLockD3D As New Object()

    ' GDI 资源
    Private _frameDataForGDI As Byte()
    Private ReadOnly _frameDataLockGDI As New Object()
    Private _currentBitmap As Bitmap
    Private _gdiBuffer As BufferedGraphics

    ' 静态着色器字节码
    Private Shared _vertexShaderBytecode As ReadOnlyMemory(Of Byte)
    Private Shared _pixelShaderBytecode As ReadOnlyMemory(Of Byte)

    ' ---------- 静态构造 ----------
    Shared Sub New()
        If LicenseManager.UsageMode = LicenseUsageMode.Designtime Then Return
        Try
            Dim vsSource = "
                struct VSInput { float3 position : POSITION; float2 texcoord : TEXCOORD; };
                struct VSOutput { float4 position : SV_POSITION; float2 texcoord : TEXCOORD; };
                VSOutput main(VSInput input) {
                    VSOutput output;
                    output.position = float4(input.position, 1);
                    output.texcoord = input.texcoord;
                    return output;
                }"
            Dim psSource = "
                Texture2D texture0 : register(t0);
                SamplerState sampler0 : register(s0);
                float4 main(float4 position : SV_POSITION, float2 texcoord : TEXCOORD) : SV_TARGET {
                    return texture0.Sample(sampler0, texcoord);
                }"
            _vertexShaderBytecode = Compiler.Compile(vsSource, "main", "vs_5_0", ShaderFlags.None)
            _pixelShaderBytecode = Compiler.Compile(psSource, "main", "ps_5_0", ShaderFlags.None)
        Catch ex As Exception
            Debug.WriteLine($"Shader compilation failed: {ex.Message}")
        End Try
    End Sub

    ' ---------- 构造函数 ----------
    Public Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint Or ControlStyles.ResizeRedraw Or ControlStyles.DoubleBuffer, True)
        BackColor = Color.Black

        _playTimer = New Timer() With {.Interval = 100}
        AddHandler _playTimer.Tick, AddressOf PlayTimer_Tick
        _renderTimer = New Timer() With {.Interval = 33}
        AddHandler _renderTimer.Tick, AddressOf RenderTimer_Tick

        AddHandler HandleCreated, AddressOf OnHandleCreated
        AddHandler HandleDestroyed, AddressOf OnHandleDestroyed
        AddHandler Resize, AddressOf VideoPanel_Resize

        ' 设计时显示占位符
        If LicenseManager.UsageMode = LicenseUsageMode.Designtime Then
            BackColor = Color.DimGray
            Using g = CreateGraphics()
                g.DrawString("Video Panel", Font, Brushes.White, 10, 10)
            End Using
        End If
    End Sub

    ' ---------- 公共属性 ----------
    <Category("Behavior"), Description("FFmpeg 可执行文件路径（支持环境变量）")>
    Public Property FFmpegPath As String
        Get
            Return _ffmpegPath
        End Get
        Set(value As String)
            _ffmpegPath = value
        End Set
    End Property

    <Category("Behavior"), Description("FFprobe 可执行文件路径")>
    Public Property FFprobePath As String
        Get
            Return _ffprobePath
        End Get
        Set(value As String)
            _ffprobePath = value
        End Set
    End Property

    <Browsable(False)>
    Public ReadOnly Property IsPlaying As Boolean
        Get
            Return _isPlaying
        End Get
    End Property

    <Browsable(False)>
    Public Property Position As TimeSpan
        Get
            Return _position
        End Get
        Set(value As TimeSpan)
            If value >= TimeSpan.Zero AndAlso value <= _duration Then Seek(value)
        End Set
    End Property

    <Browsable(False)>
    Public ReadOnly Property Duration As TimeSpan
        Get
            Return _duration
        End Get
    End Property

    <Category("Audio"), Description("音量 0.0 ~ 1.0")>
    Public Property Volume As Single
        Get
            Return _volume
        End Get
        Set(value As Single)
            _volume = Math.Clamp(value, 0.0F, 1.0F)
            If _waveOut IsNot Nothing Then _waveOut.Volume = If(_isMuted, 0.0F, _volume)
        End Set
    End Property

    <Category("Audio"), Description("静音")>
    Public Property Muted As Boolean
        Get
            Return _isMuted
        End Get
        Set(value As Boolean)
            _isMuted = value
            If _waveOut IsNot Nothing Then _waveOut.Volume = If(value, 0.0F, _volume)
        End Set
    End Property

    <Category("Rendering"), Description("渲染器类型：Auto/GDI/Direct3D11")>
    Public Property Renderer As RendererType
        Get
            Return _renderer
        End Get
        Set(value As RendererType)
            If _renderer = value Then Return
            Dim wasPlaying = _isPlaying
            Dim currentPos = _position
            If _currentFile IsNot Nothing Then
                StopPlayback()
                _renderer = value
                LoadMedia(_currentFile)
                If wasPlaying Then
                    Position = currentPos
                    Play()
                End If
            Else
                _renderer = value
            End If
            RaiseEvent RendererChanged(value)
        End Set
    End Property

    <Browsable(False)>
    Public ReadOnly Property CurrentRendererType As RendererType
        Get
            Return _renderer
        End Get
    End Property

    <Browsable(False)>
    Public ReadOnly Property CurrentFilePath As String
        Get
            Return _currentFile
        End Get
    End Property

    ' ---------- 公共方法 ----------
    Public Sub LoadMedia(filePath As String)
        If Not File.Exists(filePath) Then
            OnError($"文件不存在: {filePath}")
            Return
        End If

        StopPlayback()

        Dim info = ProbeMedia(filePath)
        If info Is Nothing Then
            OnError("无法获取媒体信息，请检查 ffprobe 路径或文件格式")
            Return
        End If

        _currentFile = filePath
        _duration = info.Duration
        _videoFPS = If(info.VideoFPS > 0, info.VideoFPS, 30.0)
        _hasVideo = info.VideoWidth > 0 AndAlso info.VideoHeight > 0

        ' 设置目标帧尺寸（保持原始比例，限制最大宽度 1280）
        Dim targetWidth = Math.Min(info.VideoWidth, 1280)
        Dim targetHeight = CInt(info.VideoHeight * (targetWidth / CDbl(info.VideoWidth)))
        If targetHeight Mod 2 = 1 Then targetHeight += 1
        _frameWidth = targetWidth
        _frameHeight = targetHeight

        _audioFormat = New WaveFormat(48000, 16, 2)
        _audioProvider = New BufferedWaveProvider(_audioFormat) With {
            .DiscardOnBufferOverflow = True,
            .BufferDuration = TimeSpan.FromSeconds(1.5)
        }
        _waveOut = New WaveOutEvent()
        _waveOut.Init(_audioProvider)
        _waveOut.Volume = If(_isMuted, 0.0F, _volume)

        ' 初始化渲染器
        If _renderer = RendererType.Auto Then
            If InitializeDirect3D11() Then
                _renderer = RendererType.Direct3D11
            Else
                _renderer = RendererType.GDI
                InitializeGDI()
            End If
        ElseIf _renderer = RendererType.Direct3D11 Then
            If Not InitializeDirect3D11() Then
                OnError("Direct3D 11 初始化失败，回退到 GDI 模式")
                _renderer = RendererType.GDI
                InitializeGDI()
            End If
        Else
            InitializeGDI()
        End If

        StartFFmpegProcesses()

        _cts = New CancellationTokenSource()
        _videoThread = New Thread(AddressOf VideoDecodeLoop) With {.IsBackground = True, .Priority = ThreadPriority.AboveNormal}
        _audioThread = New Thread(AddressOf AudioDecodeLoop) With {.IsBackground = True}
        _videoThread.Start()
        _audioThread.Start()

        _syncFrameDuration = 1.0 / _videoFPS
        _decodedFrameCount = 0
        _videoClockBase = TimeSpan.Zero
        _avSyncOffset = 0.0

        RaiseEvent MediaLoaded(info)
        Play()

        ' 确保第一帧尽快显示
        Task.Delay(100).ContinueWith(Sub() Invalidate(), TaskScheduler.FromCurrentSynchronizationContext())
    End Sub

    Public Sub Play()
        If _duration = TimeSpan.Zero OrElse _isPlaying Then Return

        If _isPaused Then
            Seek(_position, suppressPlayRestore:=True)
            _isPlaying = True
            _isPaused = False
            _waveOut?.Play()
            _playStartTime = DateTime.Now
            _playTimer.Start()
            _renderTimer.Start()
            _pauseEvent.Set()
            RaiseEvent PlayStateChanged(True)
        Else
            _playStartTime = DateTime.Now
            _basePosition = _position
            _isPlaying = True
            _waveOut?.Play()
            _playTimer.Start()
            _renderTimer.Start()
            _pauseEvent.Set()
            RaiseEvent PlayStateChanged(True)
        End If
    End Sub

    Public Sub Pause()
        If Not _isPlaying OrElse _isPaused Then Return

        Dim audioPos = GetAudioPosition()
        If audioPos > TimeSpan.Zero AndAlso audioPos <= _duration Then
            _position = audioPos
            _basePosition = _position
            RaiseEvent PositionChanged(_position)
        End If

        _isPlaying = False
        _isPaused = True
        _playTimer.Stop()
        _renderTimer.Stop()
        _waveOut?.Pause()
        _pauseEvent.Reset()
        RaiseEvent PlayStateChanged(False)
    End Sub

    Public Sub StopPlayback()
        _isPlaying = False
        _isPaused = False
        _playTimer.Stop()
        _renderTimer?.Stop()
        _playStartTime = Nothing

        _cts?.Cancel()
        _videoThread?.Join(500)
        _audioThread?.Join(500)
        _cts?.Dispose()
        _cts = Nothing

        _waveOut?.Stop()
        _waveOut?.Dispose()
        _waveOut = Nothing
        _audioProvider?.ClearBuffer()

        KillProcess(_ffmpegVideoProcess)
        KillProcess(_ffmpegAudioProcess)
        _videoStream?.Dispose()
        _audioStream?.Dispose()
        _videoStream = Nothing
        _audioStream = Nothing

        While _frameQueue.TryDequeue(Nothing) : End While
        _position = TimeSpan.Zero
        _basePosition = TimeSpan.Zero
        _decodedFrameCount = 0
    End Sub

    Public Sub Seek(target As TimeSpan, Optional suppressPlayRestore As Boolean = False)
        If _duration = TimeSpan.Zero Then Return
        target = TimeSpan.FromSeconds(Math.Clamp(target.TotalSeconds, 0, _duration.TotalSeconds))

        Dim wasPlaying = (Not suppressPlayRestore) AndAlso _isPlaying AndAlso Not _isPaused

        If wasPlaying Then
            _waveOut?.Stop()
            _playTimer.Stop()
            _renderTimer.Stop()
        End If

        _cts?.Cancel()
        _videoThread?.Join(300)
        _audioThread?.Join(300)

        While _frameQueue.TryDequeue(Nothing) : End While
        _audioProvider?.ClearBuffer()

        _position = target
        _basePosition = target
        _videoClockBase = target
        _decodedFrameCount = 0
        _playStartTime = Nothing

        RestartFFmpeg()

        _cts = New CancellationTokenSource()
        _videoThread = New Thread(AddressOf VideoDecodeLoop) With {.IsBackground = True, .Priority = ThreadPriority.AboveNormal}
        _audioThread = New Thread(AddressOf AudioDecodeLoop) With {.IsBackground = True}
        _videoThread.Start()
        _audioThread.Start()

        If wasPlaying Then
            _isPlaying = True
            _isPaused = False
            _waveOut?.Play()
            _playStartTime = DateTime.Now
            _playTimer.Start()
            _renderTimer.Start()
            _pauseEvent.Set()
            RaiseEvent PlayStateChanged(True)
        Else
            _isPlaying = False
            _isPaused = True
            _pauseEvent.Reset()
        End If

        RaiseEvent PositionChanged(_position)
    End Sub

    Public Sub ToggleFullscreen()
        Dim form = FindForm()
        If form Is Nothing Then Return

        Static windowedBounds As Rectangle
        Static isFullscreen As Boolean = False

        If Not isFullscreen Then
            windowedBounds = form.Bounds
            form.FormBorderStyle = FormBorderStyle.None
            form.WindowState = FormWindowState.Normal
            form.Bounds = Screen.PrimaryScreen.Bounds
            isFullscreen = True
        Else
            form.FormBorderStyle = FormBorderStyle.Sizable
            form.Bounds = windowedBounds
            isFullscreen = False
        End If

        If _renderer = RendererType.Direct3D11 Then
            UpdateViewportAndRenderTarget()
        Else
            Invalidate()
        End If
    End Sub

    ' ---------- 内部核心逻辑 ----------
    Private Sub StartFFmpegProcesses()
        If _hasVideo Then StartVideoProcess()
        StartAudioProcess()
    End Sub

    Private Sub StartVideoProcess()
        If Not _hasVideo Then Return
        If Not File.Exists(_ffmpegPath) Then
            OnError($"未找到 ffmpeg.exe，请设置 FFmpegPath 属性或将其放入 PATH")
            Return
        End If

        Dim args = $"-ss {_position.TotalSeconds:F3} -i ""{_currentFile}"" -map 0:v -f rawvideo -pix_fmt bgra -vf scale={_frameWidth}:{_frameHeight} -an -"
        _ffmpegVideoProcess = New Process With {
            .StartInfo = New ProcessStartInfo(_ffmpegPath, args) With {
                .UseShellExecute = False,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .CreateNoWindow = True
            }
        }
        _ffmpegVideoProcess.Start()
        _videoStream = _ffmpegVideoProcess.StandardOutput.BaseStream
        _ffmpegVideoProcess.BeginErrorReadLine()
    End Sub

    Private Sub StartAudioProcess()
        If Not File.Exists(_ffmpegPath) Then
            OnError($"未找到 ffmpeg.exe，无法播放音频")
            Return
        End If

        Dim args = $"-ss {_position.TotalSeconds:F3} -i ""{_currentFile}"" -map 0:a -f s16le -acodec pcm_s16le -ar 48000 -ac 2 -vn -"
        _ffmpegAudioProcess = New Process With {
            .StartInfo = New ProcessStartInfo(_ffmpegPath, args) With {
                .UseShellExecute = False,
                .RedirectStandardOutput = True,
                .RedirectStandardError = True,
                .CreateNoWindow = True
            }
        }
        _ffmpegAudioProcess.Start()
        _audioStream = _ffmpegAudioProcess.StandardOutput.BaseStream
        _ffmpegAudioProcess.BeginErrorReadLine()
    End Sub

    Private Sub RestartFFmpeg()
        KillProcess(_ffmpegVideoProcess)
        KillProcess(_ffmpegAudioProcess)
        _videoStream?.Dispose()
        _audioStream?.Dispose()
        _videoStream = Nothing
        _audioStream = Nothing
        StartFFmpegProcesses()
    End Sub

    Private Sub VideoDecodeLoop()
        Dim frameSize = _frameWidth * _frameHeight * 4
        If frameSize <= 0 Then Return
        Dim buffer(frameSize - 1) As Byte

        While Not _cts.Token.IsCancellationRequested
            _pauseEvent.WaitOne()

            While _frameQueue.Count >= MaxQueueSize AndAlso Not _cts.Token.IsCancellationRequested
                Thread.Sleep(1)
            End While

            Dim bytesRead = 0
            While bytesRead < frameSize
                Dim n = _videoStream.Read(buffer, bytesRead, frameSize - bytesRead)
                If n = 0 Then
                    If _ffmpegVideoProcess IsNot Nothing AndAlso _ffmpegVideoProcess.HasExited Then
                        Return
                    End If
                    Thread.Sleep(5)
                    Continue While
                End If
                bytesRead += n
            End While

            Dim framePTS = _videoClockBase + TimeSpan.FromSeconds(_decodedFrameCount * _syncFrameDuration)
            _frameQueue.Enqueue(New VideoFrame With {.Data = buffer.Clone(), .PTS = framePTS})
            _decodedFrameCount += 1
        End While
    End Sub

    Private Sub AudioDecodeLoop()
        If _audioStream Is Nothing OrElse _audioProvider Is Nothing Then Return
        Dim bufSize = 32768
        Dim buf = New Byte(bufSize - 1) {}

        While Not _cts.Token.IsCancellationRequested
            If _audioProvider.BufferedDuration.TotalSeconds > 0.5 Then
                Thread.Sleep(10)
                Continue While
            End If

            Dim n = _audioStream.Read(buf, 0, buf.Length)
            If n > 0 Then
                _audioProvider.AddSamples(buf, 0, n)
            Else
                Thread.Sleep(50)
            End If
        End While
    End Sub

    Private Sub RenderTimer_Tick(sender As Object, e As EventArgs)
        If Not _isPlaying OrElse _isPaused Then Return

        Dim audioPos = GetAudioPosition()
        Dim lastFrame As VideoFrame = Nothing

        ' 从队列中取出所有 PTS <= audioPos 的帧
        While _frameQueue.TryPeek(lastFrame) AndAlso lastFrame.PTS <= audioPos
            _frameQueue.TryDequeue(lastFrame)
        End While

        If lastFrame Is Nothing Then
            ' 没有可显示的帧，尝试重绘最后一帧
            If _frameQueue.TryPeek(lastFrame) Then
                ' 保留队列头部不取出，但渲染它
            Else
                Return
            End If
        End If

        ' 渲染
        If _renderer = RendererType.Direct3D11 Then
            UpdateFrameDataForD3D(lastFrame.Data)
            RenderD3D11()
        Else
            UpdateFrameDataForGDI(lastFrame.Data)
            Invalidate()
        End If
    End Sub

    Private Function GetAudioPosition() As TimeSpan
        If _waveOut Is Nothing OrElse _waveOut.PlaybackState = PlaybackState.Stopped Then
            Return _position
        End If
        Try
            Dim bytesPlayed = _waveOut.GetPosition()
            Dim secondsPlayed = CDbl(bytesPlayed) / _audioFormat.AverageBytesPerSecond
            ' 动态同步偏移校正
            Dim corrected = secondsPlayed - _avSyncOffset
            If corrected < 0 Then corrected = 0
            Dim pos = _basePosition + TimeSpan.FromSeconds(corrected)
            If pos > _duration Then pos = _duration
            Return pos
        Catch
            Return _position
        End Try
    End Function

    Private Sub PlayTimer_Tick(sender As Object, e As EventArgs)
        If Not _isPlaying OrElse _isPaused Then Return
        If _duration = TimeSpan.Zero Then Return

        Dim audioPos = GetAudioPosition()
        If audioPos >= _duration Then
            RaiseEvent MediaEnded()
            StopPlayback()
            Return
        End If

        If audioPos <> _position Then
            _position = audioPos
            RaiseEvent PositionChanged(_position)
        End If
    End Sub

    ' ---------- 渲染器实现 ----------
    Private Sub InitializeGDI()
        _currentBitmap?.Dispose()
        If _frameWidth > 0 AndAlso _frameHeight > 0 Then
            _currentBitmap = New Bitmap(_frameWidth, _frameHeight, Imaging.PixelFormat.Format32bppArgb)
        End If
        _gdiBuffer = BufferedGraphicsManager.Current.Allocate(CreateGraphics(), ClientRectangle)
    End Sub

    Private Sub UpdateFrameDataForGDI(frameData As Byte())
        SyncLock _frameDataLockGDI
            _frameDataForGDI = frameData.Clone()
        End SyncLock
    End Sub

    Private Sub DrawWithGDI(e As PaintEventArgs)
        Dim frameCopy As Byte() = Nothing
        SyncLock _frameDataLockGDI
            If _frameDataForGDI IsNot Nothing Then frameCopy = _frameDataForGDI.Clone()
        End SyncLock

        If frameCopy Is Nothing OrElse frameCopy.Length = 0 Then
            e.Graphics.Clear(Color.Black)
            Return
        End If

        If _currentBitmap Is Nothing OrElse _currentBitmap.Width <> _frameWidth OrElse _currentBitmap.Height <> _frameHeight Then
            _currentBitmap?.Dispose()
            _currentBitmap = New Bitmap(_frameWidth, _frameHeight, Imaging.PixelFormat.Format32bppArgb)
        End If

        Dim bmpData = _currentBitmap.LockBits(New Rectangle(0, 0, _frameWidth, _frameHeight),
                                               Imaging.ImageLockMode.WriteOnly,
                                               Imaging.PixelFormat.Format32bppArgb)
        Try
            Marshal.Copy(frameCopy, 0, bmpData.Scan0, frameCopy.Length)
        Finally
            _currentBitmap.UnlockBits(bmpData)
        End Try

        Dim clientRect = ClientRectangle
        Dim scaleW = CDbl(clientRect.Width) / _frameWidth
        Dim scaleH = CDbl(clientRect.Height) / _frameHeight
        Dim scale = Math.Min(scaleW, scaleH)
        Dim drawWidth = CInt(_frameWidth * scale)
        Dim drawHeight = CInt(_frameHeight * scale)
        Dim drawX = (clientRect.Width - drawWidth) \ 2
        Dim drawY = (clientRect.Height - drawHeight) \ 2
        e.Graphics.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        e.Graphics.DrawImage(_currentBitmap, drawX, drawY, drawWidth, drawHeight)
    End Sub

    Private Function InitializeDirect3D11() As Boolean
        SyncLock _d3dLock
            DisposeDirect3D11()
            Try
                If ClientSize.Width <= 0 OrElse ClientSize.Height <= 0 Then
                    Size = New Size(800, 600)
                End If

                Dim swapChainDesc As New SwapChainDescription()
                swapChainDesc.BufferDescription = New ModeDescription(ClientSize.Width, ClientSize.Height, New Rational(60, 1), DXGI.Format.B8G8R8A8_UNorm)
                swapChainDesc.SampleDescription = New SampleDescription(1, 0)
                swapChainDesc.BufferUsage = DXGI.Usage.RenderTargetOutput
                swapChainDesc.BufferCount = 2
                swapChainDesc.OutputWindow = Handle
                swapChainDesc.Windowed = True
                swapChainDesc.SwapEffect = DXGI.SwapEffect.Discard

                Dim featureLevels = {FeatureLevel.Level_11_0, FeatureLevel.Level_10_1, FeatureLevel.Level_10_0, FeatureLevel.Level_9_3}
                _d3dDevice = D3D11.D3D11CreateDevice(DriverType.Hardware, DeviceCreationFlags.None, featureLevels)
                If _d3dDevice Is Nothing Then
                    _d3dDevice = D3D11.D3D11CreateDevice(DriverType.Warp, DeviceCreationFlags.None, featureLevels)
                End If
                If _d3dDevice Is Nothing Then Return False

                _d3dContext = _d3dDevice.ImmediateContext

                Using dxgiDevice As IDXGIDevice = _d3dDevice.QueryInterface(Of IDXGIDevice)()
                    Using dxgiAdapter = dxgiDevice.GetAdapter()
                        Using dxgiFactory = dxgiAdapter.GetParent(Of IDXGIFactory)()
                            _swapChain = dxgiFactory.CreateSwapChain(_d3dDevice, swapChainDesc)
                        End Using
                    End Using
                End Using

                If _swapChain Is Nothing Then Return False

                Dim backBuffer = _swapChain.GetBuffer(Of ID3D11Texture2D)(0)
                _renderTargetView = _d3dDevice.CreateRenderTargetView(backBuffer)
                backBuffer.Dispose()

                Dim texDesc As New Texture2DDescription()
                texDesc.Width = _frameWidth
                texDesc.Height = _frameHeight
                texDesc.MipLevels = 1
                texDesc.ArraySize = 1
                texDesc.Format = DXGI.Format.B8G8R8A8_UNorm
                texDesc.SampleDescription = New SampleDescription(1, 0)
                texDesc.Usage = ResourceUsage.Dynamic
                texDesc.BindFlags = BindFlags.ShaderResource
                texDesc.CPUAccessFlags = CpuAccessFlags.Write
                _texture2D = _d3dDevice.CreateTexture2D(texDesc)
                If _texture2D Is Nothing Then Return False

                _shaderResourceView = _d3dDevice.CreateShaderResourceView(_texture2D)

                ' 创建着色器和顶点缓冲
                _vertexShader = _d3dDevice.CreateVertexShader(_vertexShaderBytecode.Span)
                _pixelShader = _d3dDevice.CreatePixelShader(_pixelShaderBytecode.Span)

                Dim elements() As InputElementDescription = {
                    New InputElementDescription("POSITION", 0, DXGI.Format.R32G32B32_Float, 0, 0),
                    New InputElementDescription("TEXCOORD", 0, DXGI.Format.R32G32_Float, 12, 0)
                }
                _inputLayout = _d3dDevice.CreateInputLayout(elements, _vertexShaderBytecode.ToArray())

                Dim vertices As Single() = {
                    -1, -1, 0, 0, 1,
                     1, -1, 0, 1, 1,
                    -1,  1, 0, 0, 0,
                     1,  1, 0, 1, 0
                }
                _vertexBuffer = _d3dDevice.CreateBuffer(vertices, BindFlags.VertexBuffer)

                Dim samplerDesc As New SamplerDescription()
                samplerDesc.Filter = Filter.MinMagMipLinear
                samplerDesc.AddressU = TextureAddressMode.Clamp
                samplerDesc.AddressV = TextureAddressMode.Clamp
                _samplerState = _d3dDevice.CreateSamplerState(samplerDesc)

                Dim rasterDesc As New RasterizerDescription()
                rasterDesc.CullMode = CullMode.None
                rasterDesc.FillMode = FillMode.Solid
                _rasterState = _d3dDevice.CreateRasterizerState(rasterDesc)

                UpdateViewportAndRenderTarget()
                Return True
            Catch ex As Exception
                Debug.WriteLine($"D3D11 init error: {ex.Message}")
                Return False
            End Try
        End SyncLock
    End Function

    Private Sub UpdateViewportAndRenderTarget()
        If _swapChain Is Nothing Then Return
        SyncLock _d3dLock
            _renderTargetView?.Dispose()
            _swapChain.ResizeBuffers(2, ClientSize.Width, ClientSize.Height, DXGI.Format.B8G8R8A8_UNorm, SwapChainFlags.None)
            Dim backBuffer = _swapChain.GetBuffer(Of ID3D11Texture2D)(0)
            _renderTargetView = _d3dDevice.CreateRenderTargetView(backBuffer)
            backBuffer.Dispose()
            Dim viewport As New Viewport(0, 0, ClientSize.Width, ClientSize.Height)
            _d3dContext.RSSetViewports(New Viewport() {viewport})
        End SyncLock
    End Sub

    Private Sub UpdateFrameDataForD3D(frameData As Byte())
        SyncLock _frameDataLockD3D
            _frameDataForD3D = frameData.Clone()
        End SyncLock
    End Sub

    Private Sub RenderD3D11()
        SyncLock _d3dLock
            If _d3dContext Is Nothing OrElse _renderTargetView Is Nothing Then Return

            Dim frameData As Byte() = Nothing
            SyncLock _frameDataLockD3D
                If _frameDataForD3D IsNot Nothing Then
                    frameData = _frameDataForD3D
                    _frameDataForD3D = Nothing
                End If
            End SyncLock

            If frameData IsNot Nothing AndAlso _texture2D IsNot Nothing Then
                Dim mapped = _d3dContext.Map(_texture2D, 0, MapMode.WriteDiscard)
                If mapped.DataPointer <> IntPtr.Zero Then
                    Dim srcPitch = _frameWidth * 4
                    If mapped.RowPitch = srcPitch Then
                        Marshal.Copy(frameData, 0, mapped.DataPointer, frameData.Length)
                    Else
                        For y = 0 To _frameHeight - 1
                            Dim srcOffset = y * srcPitch
                            Dim dstOffset = y * mapped.RowPitch
                            Marshal.Copy(frameData, srcOffset, IntPtr.Add(mapped.DataPointer, dstOffset), srcPitch)
                        Next
                    End If
                    _d3dContext.Unmap(_texture2D, 0)
                End If
            End If

            _d3dContext.OMSetRenderTargets(_renderTargetView)
            _d3dContext.RSSetState(_rasterState)
            _d3dContext.IASetVertexBuffers(0, New ID3D11Buffer() {_vertexBuffer}, New UInteger() {20UI}, New UInteger() {0UI})
            _d3dContext.IASetInputLayout(_inputLayout)
            _d3dContext.IASetPrimitiveTopology(PrimitiveTopology.TriangleStrip)
            _d3dContext.VSSetShader(_vertexShader)
            _d3dContext.PSSetShader(_pixelShader)
            _d3dContext.PSSetShaderResources(0, New ID3D11ShaderResourceView() {_shaderResourceView})
            _d3dContext.PSSetSamplers(0, New ID3D11SamplerState() {_samplerState})
            _d3dContext.Draw(4, 0)
            _swapChain.Present(0, 0)
        End SyncLock
    End Sub

    Private Sub DisposeDirect3D11()
        SyncLock _d3dLock
            _rasterState?.Dispose()
            _samplerState?.Dispose()
            _vertexBuffer?.Dispose()
            _inputLayout?.Dispose()
            _pixelShader?.Dispose()
            _vertexShader?.Dispose()
            _shaderResourceView?.Dispose()
            _texture2D?.Dispose()
            _renderTargetView?.Dispose()
            _swapChain?.Dispose()
            _d3dContext?.Dispose()
            _d3dDevice?.Dispose()
            _rasterState = Nothing
            _samplerState = Nothing
            _vertexBuffer = Nothing
            _inputLayout = Nothing
            _pixelShader = Nothing
            _vertexShader = Nothing
            _shaderResourceView = Nothing
            _texture2D = Nothing
            _renderTargetView = Nothing
            _swapChain = Nothing
            _d3dContext = Nothing
            _d3dDevice = Nothing
        End SyncLock
    End Sub

    ' ---------- 辅助工具 ----------
    Private Function ProbeMedia(file As String) As MediaInfo
        Try
            If Not File.Exists(_ffprobePath) Then
                OnError($"ffprobe.exe 未找到: {_ffprobePath}")
                Return Nothing
            End If

            Using p As New Process()
                p.StartInfo.FileName = _ffprobePath
                p.StartInfo.Arguments = $"-v error -print_format json -show_format -show_streams ""{file}"""
                p.StartInfo.UseShellExecute = False
                p.StartInfo.RedirectStandardOutput = True
                p.StartInfo.StandardOutputEncoding = Encoding.UTF8
                p.StartInfo.CreateNoWindow = True
                p.Start()
                Dim output = p.StandardOutput.ReadToEnd()
                p.WaitForExit()

                Dim root = JObject.Parse(output)
                Dim streams = root("streams")
                Dim fmt = root("format")

                Dim v = streams?.FirstOrDefault(Function(s) s("codec_type")?.ToString() = "video")
                Dim a = streams?.FirstOrDefault(Function(s) s("codec_type")?.ToString() = "audio")

                Dim durationSeconds = If(fmt("duration") IsNot Nothing, CDbl(fmt("duration").ToString()), 0.0)

                Dim fps = 0.0
                If v IsNot Nothing Then
                    Dim avgFrameRate = v("avg_frame_rate")?.ToString()
                    If Not String.IsNullOrEmpty(avgFrameRate) AndAlso avgFrameRate.Contains("/") Then
                        Dim parts = avgFrameRate.Split("/"c)
                        If parts.Length = 2 Then
                            Dim num = CDbl(parts(0))
                            Dim den = CDbl(parts(1))
                            If den > 0 Then fps = num / den
                        End If
                    End If
                End If

                Dim artist = "", album = ""
                If fmt("tags") IsNot Nothing Then
                    Dim tags = fmt("tags")
                    If tags("artist") IsNot Nothing Then artist = tags("artist").ToString()
                    If tags("album") IsNot Nothing Then album = tags("album").ToString()
                End If

                Return New MediaInfo With {
                    .FilePath = file,
                    .Duration = TimeSpan.FromSeconds(durationSeconds),
                    .VideoWidth = If(v IsNot Nothing, v("width")?.Value(Of Integer)(), 0),
                    .VideoHeight = If(v IsNot Nothing, v("height")?.Value(Of Integer)(), 0),
                    .VideoFPS = fps,
                    .AudioSampleRate = If(a IsNot Nothing AndAlso a("sample_rate") IsNot Nothing, Integer.Parse(a("sample_rate").ToString()), 48000),
                    .AudioChannels = If(a IsNot Nothing, a("channels")?.Value(Of Integer)(), 2),
                    .Artist = artist,
                    .Album = album
                }
            End Using
        Catch ex As Exception
            Debug.WriteLine($"Probe error: {ex.Message}")
            Return Nothing
        End Try
    End Function

    Private Sub KillProcess(ByRef proc As Process)
        If proc IsNot Nothing Then
            Try
                If Not proc.HasExited Then
                    proc.StandardInput.WriteLine("q")
                    proc.WaitForExit(1000)
                    proc.Kill()
                End If
            Catch
            End Try
            proc.Close()
            proc = Nothing
        End If
    End Sub

    Private Sub OnError(msg As String)
        If InvokeRequired Then
            Invoke(Sub() RaiseEvent ErrorOccurred(msg))
        Else
            RaiseEvent ErrorOccurred(msg)
        End If
    End Sub

    ' ---------- 事件处理 ----------
    Private Sub OnHandleCreated(sender As Object, e As EventArgs)
        If _renderer = RendererType.Direct3D11 AndAlso _frameWidth > 0 Then
            InitializeDirect3D11()
        End If
    End Sub

    Private Sub OnHandleDestroyed(sender As Object, e As EventArgs)
        DisposeDirect3D11()
    End Sub

    Private Sub VideoPanel_Resize(sender As Object, e As EventArgs)
        If _renderer = RendererType.Direct3D11 Then
            UpdateViewportAndRenderTarget()
        Else
            Invalidate()
        End If
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        If LicenseManager.UsageMode = LicenseUsageMode.Designtime Then
            MyBase.OnPaint(e)
            Return
        End If

        If _renderer = RendererType.GDI Then
            DrawWithGDI(e)
        Else
            MyBase.OnPaint(e)
        End If
    End Sub

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            StopPlayback()
            _playTimer?.Dispose()
            _renderTimer?.Dispose()
            _waveOut?.Dispose()
            _currentBitmap?.Dispose()
            _gdiBuffer?.Dispose()
            DisposeDirect3D11()
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class