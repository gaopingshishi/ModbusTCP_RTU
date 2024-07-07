using System.IO.Ports;
using System.Net;
using System.Net.Sockets;
using ModbusTCP_RTU.Exceptions;

namespace ModbusTCP_RTU
{
    #region Modbus 协议信息类

    /// <summary>
    /// Modbus 信息类。
    /// </summary>
    public class ModbusProtocol
    {
        public enum ProtocolType
        {
            ModbusTcp = 0,
            ModbusUdp = 1,
            ModbusRtu = 2
        }

        public DateTime TimeStamp;
        public bool Request;
        public bool Response;

        /// <summary>
        /// 事务标识符(字节形式)（仅TCP模式）
        /// </summary>
        public ushort TransactionIdentifier;

        /// <summary>
        /// 协议标识符 0x0000 标识ModbusTCP协议（仅TCP模式）
        /// </summary>
        public ushort ProtocolIdentifier;

        /// <summary>
        /// 从当前位置开始，后边的报文长度（仅TCP模式）
        /// </summary>
        public ushort Length;

        /// <summary>
        /// 获取或设置串行连接时的单位标识符 / 从站的地址（默认 = 0）
        /// </summary>
        public byte UnitIdentifier;

        /// <summary>
        /// 功能码
        /// </summary>
        public byte FunctionCode;

        /// <summary>
        /// 寄存器开始地址
        /// </summary>
        public ushort StartingAddress;

        /// <summary>
        /// 读取的起始地址（仅 读写多个寄存器 功能用到）
        /// </summary>
        public ushort StartingAddressRead;

        /// <summary>
        /// 写入的起始地址（仅 读写多个寄存器 功能用到）
        /// </summary>
        public ushort StartingAddressWrite;

        /// <summary>
        /// 数量（可能是寄存器，也可能是字节或线圈，以实际功能为准）
        /// </summary>
        public ushort Quantity;

        /// <summary>
        /// 读取数量（仅 读写多个寄存器 功能用到）
        /// </summary>
        public ushort QuantityRead;

        /// <summary>
        /// 写入数量（仅 读写多个寄存器 功能用到）
        /// </summary>
        public ushort QuantityWrite;

        /// <summary>
        /// 字节数量
        /// </summary>
        public byte ByteCount;

        /// <summary>
        /// 异常码
        /// </summary>
        public byte ExceptionCode;

        /// <summary>
        /// 错误标识（0x80 + 功能码）
        /// </summary>
        public byte ErrorCode;

        public ushort[]? ReceiveCoilValues;
        public ushort[]? ReceiveRegisterValues;
        public short[]? SendRegisterValues;
        public bool[]? SendCoilValues;
        public ushort Crc;
    }

    #endregion

    #region structs

    internal struct NetworkConnectionParameter
    {
        public NetworkStream NetStream; //仅适用于 TCP 连接
        public byte[] Bytes;
        public int PortIn; //仅适用于 UDP 连接
        public IPAddress? IpAddressIn; //仅适用于 UDP 连接
    }

    #endregion

    #region Tcp操作类

    internal class TcpHandler
    {
        #region 事件

        public delegate void DataChangedDelegate(object networkConnectionParameter);

        /// <summary>
        /// 数据发生变化时触发的事件。
        /// </summary>
        public event DataChangedDelegate? DataChangedEvent;

        public delegate void NumberOfClientsChangedDelegate();

        /// <summary>
        /// 连接的客户端数量发生变化时触发的事件。
        /// </summary>
        public event NumberOfClientsChangedDelegate? NumberOfClientsChangedEvent;

        #endregion

        /// <summary>
        /// TCP客户端对象
        /// </summary>
        private readonly TcpListener? _server;

        /// <summary>
        /// 已连接的客户端列表。
        /// </summary>
        private readonly List<Client> _tcpClientLastRequestList = [];

        /// <summary>
        /// 当前连接的客户端数量
        /// </summary>
        public int NumberOfConnectedClients { get; private set; }

        /// <summary>
        /// 当创建服务器 TCP 时，将侦听此 IP 地址。
        /// </summary>
        private readonly IPAddress _localIpAddress = IPAddress.Any;

        /// <summary>
        /// IP地址白名单，如果不为空，则只有白名单中的IP地址才能连接到服务器。
        /// </summary>
        private readonly string[]? _ipWhitelisting;

        /// <summary>
        /// 超时时间，如果客户端在此时间内没有发送数据，则断开连接。 单位：ms ，设置为 0 代表不启用超时。
        /// </summary>
        /// <returns></returns>
        private readonly long _timeout;

        /// <summary>
        /// 监听所有IP地址。
        /// </summary>
        /// <param name="port">要监听的 TCP 端口</param>
        /// <param name="timeout">超时时间，如果客户端在此时间内没有发送数据，则断开连接。 单位：ms ，设置为 0 代表不启用超时。</param>
        /// <param name="ipWhitelisting">IP白名单，为空代表不启用</param>
        public TcpHandler(int port, long timeout = 0, string[]? ipWhitelisting = null)
        {
            _timeout = timeout;
            _ipWhitelisting = ipWhitelisting;
            _server = new TcpListener(_localIpAddress, port);
            _server.Start();
            _server.BeginAcceptTcpClient(AcceptTcpClientCallback, null);
        }

        /// <summary>
        /// 监听特定IP地址
        /// </summary>
        /// <param name="localIpAddress">要监听的网络接口 IP 地址</param>
        /// <param name="port">要监听的 TCP 端口</param>
        /// <param name="timeout">超时时间，如果客户端在此时间内没有发送数据，则断开连接。 单位：ms ，设置为 0 代表不启用超时。</param>
        /// <param name="ipWhitelisting">IP白名单，为空代表不启用</param>
        public TcpHandler(IPAddress localIpAddress, int port, long timeout = 0, string[]? ipWhitelisting = null)
        {
            _timeout = timeout;
            _ipWhitelisting = ipWhitelisting;
            _localIpAddress = localIpAddress;
            _server = new TcpListener(_localIpAddress, port);
            _server.Start();
            _server.BeginAcceptTcpClient(AcceptTcpClientCallback, null);
        }


        /// <summary>
        /// 客户端接入后的回调函数
        /// </summary>
        /// <param name="asyncResult"></param>
        private void AcceptTcpClientCallback(IAsyncResult asyncResult)
        {
            if (_server == null)
                return;

            TcpClient? tcpClient = null;

            #region 白名单判断

            try
            {
                tcpClient = _server.EndAcceptTcpClient(asyncResult);
                tcpClient.ReceiveTimeout = 4000;
                if (_ipWhitelisting?.Length > 0)
                {
                    var ipEndpoint = tcpClient.Client.RemoteEndPoint?.ToString();
                    if (!string.IsNullOrEmpty(ipEndpoint))
                    {
                        ipEndpoint = ipEndpoint.Split(':')[0];
                        if (!_ipWhitelisting.Contains(ipEndpoint))
                        {
                            tcpClient.Client.Disconnect(false);
                            return;
                        }
                    }
                }
            }
            catch (Exception)
            {
                // ignored
            }

            #endregion

            try
            {
                //继续接受下一个客户端
                _server.BeginAcceptTcpClient(AcceptTcpClientCallback, null);

                //开始异步读取数据
                if (tcpClient == null)
                    return;
                var client = new Client(tcpClient);
                var networkStream = client.NetworkStream;
                networkStream.ReadTimeout = 4000;
                networkStream.BeginRead(client.Buffer, 0, client.Buffer.Length, ReadCallback, client);
            }
            catch (Exception)
            {
                // ignored
            }
        }

        /// <summary>
        /// 读取到数据后的回调函数
        /// </summary>
        /// <param name="asyncResult"></param>
        private void ReadCallback(IAsyncResult asyncResult)
        {
            var networkConnectionParameter = new NetworkConnectionParameter();
            if (asyncResult.AsyncState is not Client client)
                return;
            client.Ticks = DateTime.Now.Ticks;
            //通知客户端数量变化
            NumberOfConnectedClients = GetAndCleanNumberOfConnectedClients(client);
            NumberOfClientsChangedEvent?.Invoke();

            #region 触发数据处理逻辑 DataChangedEvent

            int read;
            NetworkStream? networkStream;
            try
            {
                networkStream = client.NetworkStream;
                read = networkStream.EndRead(asyncResult);
            }
            catch (Exception)
            {
                return;
            }

            if (read == 0)
                return;

            var data = new byte[read];
            Buffer.BlockCopy(client.Buffer, 0, data, 0, read);
            networkConnectionParameter.Bytes = data;
            networkConnectionParameter.NetStream = networkStream;
            DataChangedEvent?.Invoke(networkConnectionParameter);

            #endregion

            try
            {
                //继续读取数据
                networkStream.BeginRead(client.Buffer, 0, client.Buffer.Length, ReadCallback, client);
            }
            catch (Exception)
            {
                // ignored
            }
        }

        private int GetAndCleanNumberOfConnectedClients(Client client)
        {
            lock (this)
            {
                //检查是否已经在缓存列表中
                var objetExists = false;
                foreach (var _ in _tcpClientLastRequestList.Where(client.Equals))
                {
                    objetExists = true;
                }

                //关闭所有超时的客户端
                try
                {
                    if (_timeout > 0)
                    {
                        var clients = _tcpClientLastRequestList.FindAll(c =>
                            DateTime.Now.Ticks - c.Ticks > TimeSpan.TicksPerMillisecond * _timeout);
                        Parallel.ForEach(clients, c =>
                        {
                            try
                            {
                                c.NetworkStream.Close(00);
                                _tcpClientLastRequestList.Remove(c);
                            }
                            catch (Exception)
                            {
                                // ignored
                            }
                        });
                    }
                }
                catch (Exception)
                {
                    // ignored
                }

                if (!objetExists)
                    _tcpClientLastRequestList.Add(client);

                return _tcpClientLastRequestList.Count;
            }
        }

        public void Disconnect()
        {
            try
            {
                foreach (var clientLoop in _tcpClientLastRequestList)
                {
                    clientLoop.NetworkStream.Close(00);
                }
            }
            catch (Exception)
            {
                // ignored
            }

            _server?.Stop();
        }


        private class Client
        {
            /// <summary>
            /// 最近一次通讯的时刻
            /// </summary>
            public long Ticks { get; set; }

            public Client(TcpClient tcpClient)
            {
                TcpClient = tcpClient;
                var bufferSize = tcpClient.ReceiveBufferSize;
                Buffer = new byte[bufferSize];
            }

            // ReSharper disable once MemberCanBePrivate.Local
            public TcpClient TcpClient { get; }

            public byte[] Buffer { get; }

            public NetworkStream NetworkStream => TcpClient.GetStream();
        }
    }

    #endregion

    /// <summary>
    /// Modbus 服务端 从站
    /// </summary>
    public class ModbusServer
    {
        #region events

        public delegate void CoilsChangedHandler(int coil, int numberOfCoils);

        public event CoilsChangedHandler? CoilsChanged;

        public delegate void HoldingRegistersChangedHandler(int register, int numberOfRegisters);

        public event HoldingRegistersChangedHandler? HoldingRegistersChanged;

        public delegate void NumberOfConnectedClientsChangedHandler();

        public event NumberOfConnectedClientsChangedHandler? NumberOfConnectedClientsChanged;

        public delegate void LogDataChangedHandler();

        public event LogDataChangedHandler? LogDataChanged;

        #endregion

        private byte[] _bytes = new byte[2100];

        // ReSharper disable once FieldCanBeMadeReadOnly.Global
        // ReSharper disable once MemberCanBePrivate.Global
        public HoldingRegisters HoldingRegistersInfo;

        // ReSharper disable once FieldCanBeMadeReadOnly.Global
        // ReSharper disable once MemberCanBePrivate.Global
        public InputRegisters InputRegistersInfo;

        // ReSharper disable once FieldCanBeMadeReadOnly.Global
        // ReSharper disable once MemberCanBePrivate.Global
        public Coils CoilsInfo;

        // ReSharper disable once FieldCanBeMadeReadOnly.Global
        // ReSharper disable once MemberCanBePrivate.Global
        public DiscreteInputs DiscreteInputsInfo;

        /// <summary>
        /// RTU 串口名称
        /// </summary>
        private string? _serialPortName = "COM1";

        /// <summary>
        /// RTU 串口对象
        /// </summary>
        private SerialPort? _serialPort;

        /// <summary>
        /// UDP 模式中，客户端端口
        /// </summary>
        private int _portIn;

        /// <summary>
        /// UDP 模式中，客户端IP地址
        /// </summary>
        private IPAddress? _ipAddressIn;

        /// <summary>
        /// UDP 客户端对象
        /// </summary>
        private UdpClient? _udpClient;

        /// <summary>
        /// UDP 模式中，客户端IP和端口信息
        /// </summary>
        private IPEndPoint? _iPEndPoint;

        /// <summary>
        /// TCP 客户端对象
        /// </summary>
        private TcpHandler? _tcpHandler;

        /// <summary>
        /// 侦听线程
        /// </summary>
        private Thread? _listenerThread;

        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        public bool FunctionCode1Disabled { get; set; }

        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        public bool FunctionCode2Disabled { get; set; }

        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        public bool FunctionCode3Disabled { get; set; }

        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        public bool FunctionCode4Disabled { get; set; }

        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        public bool FunctionCode5Disabled { get; set; }

        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        public bool FunctionCode6Disabled { get; set; }

        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        public bool FunctionCode15Disabled { get; set; }

        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        public bool FunctionCode16Disabled { get; set; }

        // ReSharper disable once UnusedAutoPropertyAccessor.Global
        public bool FunctionCode23Disabled { get; set; }

        /// <summary>
        /// 串口变更标识，用于判断是否需要重新创建串口对象
        /// </summary>
        private bool _portChanged;

        private readonly object _lockCoils = new object();
        private readonly object _lockHoldingRegisters = new object();
        private volatile bool _shouldStop;

        private IPAddress _localIpAddress = IPAddress.Any;

        /// <summary>
        /// 创建 TCP 或 UDP 套接字时，要绑定的本地 IP 地址。
        /// </summary>
        public IPAddress LocalIpAddress
        {
            get => _localIpAddress;
            set
            {
                if (_listenerThread == null)
                    _localIpAddress = value;
            }
        }

        public ModbusServer(bool serialFlag, UdpClient udpClient)
        {
            SerialFlag = serialFlag;
            _udpClient = udpClient;
            HoldingRegistersInfo = new HoldingRegisters(this);
            InputRegistersInfo = new InputRegisters(this);
            CoilsInfo = new Coils(this);
            DiscreteInputsInfo = new DiscreteInputs(this);
        }


        public void Listen()
        {
            _listenerThread = new Thread(ListenerThread);
            _listenerThread.Start();
        }

        public void StopListening()
        {
            if (SerialFlag & (_serialPort != null))
            {
                if (_serialPort is { IsOpen: true })
                    _serialPort.Close();
                _shouldStop = true;
            }

            try
            {
                _tcpHandler?.Disconnect();
#pragma warning disable SYSLIB0006
                _listenerThread?.Abort();
#pragma warning restore SYSLIB0006
            }
            catch (Exception)
            {
                // ignored
            }

            _listenerThread?.Join();
        }

        private void ListenerThread()
        {
            if (!UdpFlag & !SerialFlag)
            {
                try
                {
                    _udpClient?.Close();
                }
                catch (Exception)
                {
                    // ignored
                }

                _tcpHandler = new TcpHandler(LocalIpAddress, Port);
                _tcpHandler.DataChangedEvent += ProcessReceivedData;
                _tcpHandler.NumberOfClientsChangedEvent += NumberOfClientsChanged;
            }
            else if (SerialFlag)
            {
                if (_serialPort != null)
                    return;
                _serialPort = new SerialPort();
                _serialPort.PortName = _serialPortName;
                _serialPort.BaudRate = BaudRate;
                _serialPort.Parity = Parity;
                _serialPort.StopBits = StopBits;
                _serialPort.WriteTimeout = 10000;
                _serialPort.ReadTimeout = 1000;
                _serialPort.DataReceived += DataReceivedHandler;
                _serialPort.Open();
            }
            else
                while (!_shouldStop)
                {
                    if (!UdpFlag)
                        continue;
                    if (_udpClient == null | _portChanged)
                    {
                        var localEndpoint = new IPEndPoint(LocalIpAddress, Port);
                        _udpClient = new UdpClient(localEndpoint);

                        _udpClient.Client.ReceiveTimeout = 1000;
                        _iPEndPoint = new IPEndPoint(IPAddress.Any, Port);
                        _portChanged = false;
                    }

                    _tcpHandler?.Disconnect();

                    try
                    {
                        if (_udpClient != null)
                        {
                            _bytes = _udpClient.Receive(ref _iPEndPoint);
                            _portIn = _iPEndPoint.Port;
                            var networkConnectionParameter = new NetworkConnectionParameter
                            {
                                Bytes = _bytes,
                                PortIn = _portIn,
                                IpAddressIn = _ipAddressIn
                            };
                            _ipAddressIn = _iPEndPoint.Address;
                            ParameterizedThreadStart pts = ProcessReceivedData;
                            var processDataThread = new Thread(pts);
                            processDataThread.Start(networkConnectionParameter);
                        }
                    }
                    catch (Exception)
                    {
                        // ignored
                    }
                }
        }

        #region 串口通讯时的数据处理

        /// <summary>
        /// 读取缓冲区
        /// </summary>
        private readonly byte[] _readBuffer = new byte[2094];

        /// <summary>
        /// 上一次数据通讯时间
        /// </summary>
        private DateTime _lastReceive;

        /// <summary>
        /// 下一个数据的位置
        /// </summary>
        private int _nextSign;

        /// <summary>
        /// RTU模式下，检查接受到的数据是否正确。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DataReceivedHandler(object sender,
            SerialDataReceivedEventArgs e)
        {
            /*
             * ModbusRTU协议中，需要用时间间隔来判断一帧报文的开始和结束，协议规定的时间为3.5个字符周期。
             * 在一帧报文开始前，必须有大于3.5个字符周期的空闲时间，一帧报文结束后，也必须要有3.5个字符周期的空闲时间，否则就会出现粘包的请况。
             * 3.5个字符周期是一个具体时间，与波特率有关
             *
             * 字符周期计算公式：字符周期 = （1 / 波特率） * 每字符位数。
             * 例如：波特率为9600 bps，一个字符由 1 个起始位、8 个数据位、1 个奇偶校验位和 1 个停止位组成，共 11 位，字符周期为（1 / 9600） * 11 = 1.145833ms
             */
            var silence = 4000 / BaudRate;
            if ((DateTime.Now.Ticks - _lastReceive.Ticks) > TimeSpan.TicksPerMillisecond * silence)
                _nextSign = 0;

            var sp = (SerialPort)sender;

            var 字节数量 = sp.BytesToRead;
            var rxByteArray = new byte[字节数量];

            sp.Read(rxByteArray, 0, 字节数量);

            Array.Copy(rxByteArray, 0, _readBuffer, _nextSign, rxByteArray.Length);
            _lastReceive = DateTime.Now;
            _nextSign = 字节数量 + _nextSign;

            if (!ModbusClient.DetectValidModbusFrame(_readBuffer, _nextSign))
                return;

            _nextSign = 0;

            var networkConnectionParameter = new NetworkConnectionParameter
            {
                Bytes = _readBuffer
            };
            ParameterizedThreadStart pts = ProcessReceivedData;
            var processDataThread = new Thread(pts);
            processDataThread.Start(networkConnectionParameter);
        }

        #endregion

        #region 客户端连接数量变更 类

        private void NumberOfClientsChanged()
        {
            if (_tcpHandler != null)
                NumberOfConnections = _tcpHandler.NumberOfConnectedClients;
            NumberOfConnectedClientsChanged?.Invoke();
        }

        #endregion

        private readonly object _lockProcessReceivedData = new();

        #region 数据处理方法

        /// <summary>
        /// 处理接收到的数据事件
        /// </summary>
        /// <param name="networkConnectionParameter"></param>
        private void ProcessReceivedData(object? networkConnectionParameter)
        {
            lock (_lockProcessReceivedData)
            {
                if (networkConnectionParameter is not NetworkConnectionParameter parameter)
                    return;

                var bytes = new byte[parameter.Bytes.Length];
                var stream = parameter.NetStream;
                var portIn = parameter.PortIn;
                var ipAddressIn = parameter.IpAddressIn;

                Array.Copy(parameter.Bytes, 0, bytes, 0,
                    parameter.Bytes.Length);

                var receiveDataThread = new ModbusProtocol();
                var sendDataThread = new ModbusProtocol();

                try
                {
                    //寄存器结构
                    var wordData = new ushort[1];
                    //字节结构
                    var byteData = new byte[2];

                    receiveDataThread.TimeStamp = DateTime.Now;
                    receiveDataThread.Request = true;

                    if (!SerialFlag)
                    {
                        //事务标识符
                        byteData[1] = bytes[0];
                        byteData[0] = bytes[1];
                        Buffer.BlockCopy(byteData, 0, wordData, 0, 2);
                        receiveDataThread.TransactionIdentifier = wordData[0];

                        //协议标识符
                        byteData[1] = bytes[2];
                        byteData[0] = bytes[3];
                        Buffer.BlockCopy(byteData, 0, wordData, 0, 2);
                        receiveDataThread.ProtocolIdentifier = wordData[0];

                        //报文长度
                        byteData[1] = bytes[4];
                        byteData[0] = bytes[5];
                        Buffer.BlockCopy(byteData, 0, wordData, 0, 2);
                        receiveDataThread.Length = wordData[0];
                    }

                    //单元标识符
                    receiveDataThread.UnitIdentifier = bytes[6 - 6 * Convert.ToInt32(SerialFlag)];
                    //检查单元标识符是否正确（即：是不是要和我进行通讯）
                    if ((receiveDataThread.UnitIdentifier != UnitIdentifier) &
                        (receiveDataThread.UnitIdentifier != 0))
                        return;

                    //功能码
                    receiveDataThread.FunctionCode = bytes[7 - 6 * Convert.ToInt32(SerialFlag)];

                    //寄存器开始地址
                    byteData[1] = bytes[8 - 6 * Convert.ToInt32(SerialFlag)];
                    byteData[0] = bytes[9 - 6 * Convert.ToInt32(SerialFlag)];
                    Buffer.BlockCopy(byteData, 0, wordData, 0, 2);
                    receiveDataThread.StartingAddress = wordData[0];

                    switch (receiveDataThread.FunctionCode)
                    {
                        case <= 4:
                            // 查询数量
                            byteData[1] = bytes[10 - 6 * Convert.ToInt32(SerialFlag)];
                            byteData[0] = bytes[11 - 6 * Convert.ToInt32(SerialFlag)];
                            Buffer.BlockCopy(byteData, 0, wordData, 0, 2);
                            receiveDataThread.Quantity = wordData[0];
                            break;
                        case 5:
                            //写单个线圈
                            receiveDataThread.ReceiveCoilValues = new ushort[1];
                            //线圈数值
                            byteData[1] = bytes[10 - 6 * Convert.ToInt32(SerialFlag)];
                            byteData[0] = bytes[11 - 6 * Convert.ToInt32(SerialFlag)];
                            Buffer.BlockCopy(byteData, 0, receiveDataThread.ReceiveCoilValues, 0, 2);
                            break;
                        case 6:
                            //写单个保持寄存器
                            receiveDataThread.ReceiveRegisterValues = new ushort[1];
                            //寄存器数值
                            byteData[1] = bytes[10 - 6 * Convert.ToInt32(SerialFlag)];
                            byteData[0] = bytes[11 - 6 * Convert.ToInt32(SerialFlag)];
                            Buffer.BlockCopy(byteData, 0, receiveDataThread.ReceiveRegisterValues, 0, 2);
                            break;
                        case 15:
                        {
                            //写多个线圈

                            //写入数量
                            byteData[1] = bytes[10 - 6 * Convert.ToInt32(SerialFlag)];
                            byteData[0] = bytes[11 - 6 * Convert.ToInt32(SerialFlag)];
                            Buffer.BlockCopy(byteData, 0, wordData, 0, 2);
                            receiveDataThread.Quantity = wordData[0];

                            //字节数量（byte类型，线圈是bit类型的，所以1个字节能存放8个线圈数据）
                            receiveDataThread.ByteCount = bytes[12 - 6 * Convert.ToInt32(SerialFlag)];

                            receiveDataThread.ReceiveCoilValues = receiveDataThread.ByteCount % 2 != 0
                                ? new ushort[receiveDataThread.ByteCount / 2 + 1]
                                : new ushort[receiveDataThread.ByteCount / 2];

                            Buffer.BlockCopy(bytes, 13 - 6 * Convert.ToInt32(SerialFlag),
                                receiveDataThread.ReceiveCoilValues, 0, receiveDataThread.ByteCount);
                            break;
                        }
                        case 16:
                        {
                            //写多个保持寄存器
                            //寄存器数量
                            byteData[1] = bytes[10 - 6 * Convert.ToInt32(SerialFlag)];
                            byteData[0] = bytes[11 - 6 * Convert.ToInt32(SerialFlag)];
                            Buffer.BlockCopy(byteData, 0, wordData, 0, 2);
                            receiveDataThread.Quantity = wordData[0];

                            //字节数量（byte类型，寄存器是short类型的，所以需要2个字节才能存放1个寄存器）
                            receiveDataThread.ByteCount = bytes[12 - 6 * Convert.ToInt32(SerialFlag)];

                            receiveDataThread.ReceiveRegisterValues = new ushort[receiveDataThread.Quantity];
                            for (var i = 0; i < receiveDataThread.Quantity; i++)
                            {
                                byteData[1] = bytes[13 + i * 2 - 6 * Convert.ToInt32(SerialFlag)];
                                byteData[0] = bytes[14 + i * 2 - 6 * Convert.ToInt32(SerialFlag)];
                                Buffer.BlockCopy(byteData, 0, receiveDataThread.ReceiveRegisterValues, i * 2, 2);
                            }

                            break;
                        }
                        case 23:
                        {
                            // 读写多个寄存器。先写后读

                            // 读取操作的第一个寄存器地址
                            byteData[1] = bytes[8 - 6 * Convert.ToInt32(SerialFlag)];
                            byteData[0] = bytes[9 - 6 * Convert.ToInt32(SerialFlag)];
                            Buffer.BlockCopy(byteData, 0, wordData, 0, 2);
                            receiveDataThread.StartingAddressRead = wordData[0];

                            // 读取的寄存器数量
                            byteData[1] = bytes[10 - 6 * Convert.ToInt32(SerialFlag)];
                            byteData[0] = bytes[11 - 6 * Convert.ToInt32(SerialFlag)];
                            Buffer.BlockCopy(byteData, 0, wordData, 0, 2);
                            receiveDataThread.QuantityRead = wordData[0];

                            // 写入操作的第一个寄存器地址
                            byteData[1] = bytes[12 - 6 * Convert.ToInt32(SerialFlag)];
                            byteData[0] = bytes[13 - 6 * Convert.ToInt32(SerialFlag)];
                            Buffer.BlockCopy(byteData, 0, wordData, 0, 2);
                            receiveDataThread.StartingAddressWrite = wordData[0];

                            // 写入的寄存器数量
                            byteData[1] = bytes[14 - 6 * Convert.ToInt32(SerialFlag)];
                            byteData[0] = bytes[15 - 6 * Convert.ToInt32(SerialFlag)];
                            Buffer.BlockCopy(byteData, 0, wordData, 0, 2);
                            receiveDataThread.QuantityWrite = wordData[0];

                            //字节数量
                            receiveDataThread.ByteCount = bytes[16 - 6 * Convert.ToInt32(SerialFlag)];
                            receiveDataThread.ReceiveRegisterValues = new ushort[receiveDataThread.QuantityWrite];
                            for (var i = 0; i < receiveDataThread.QuantityWrite; i++)
                            {
                                byteData[1] = bytes[17 + i * 2 - 6 * Convert.ToInt32(SerialFlag)];
                                byteData[0] = bytes[18 + i * 2 - 6 * Convert.ToInt32(SerialFlag)];
                                Buffer.BlockCopy(byteData, 0, receiveDataThread.ReceiveRegisterValues, i * 2, 2);
                            }

                            break;
                        }
                    }
                }
                catch (Exception)
                {
                    // ignored
                }

                CreateAnswer(receiveDataThread, sendDataThread, stream, portIn, ipAddressIn);
                CreateLogData(receiveDataThread, sendDataThread);

                LogDataChanged?.Invoke();
            }
        }

        #endregion

        #region 应答 方法

        /// <summary>
        /// 创建应答
        /// </summary>
        /// <param name="receiveData">接收数据</param>
        /// <param name="sendData">发送数据</param>
        /// <param name="stream">网络流</param>
        /// <param name="portIn">客户端端口</param>
        /// <param name="ipAddressIn">客户端IP</param>
        private void CreateAnswer(ModbusProtocol receiveData, ModbusProtocol sendData, NetworkStream stream, int portIn,
            IPAddress? ipAddressIn)
        {
            switch (receiveData.FunctionCode)
            {
                //读输出线圈
                case 1:
                    if (!FunctionCode1Disabled)
                        ReadCoils(receiveData, sendData, stream, portIn, ipAddressIn);
                    else
                    {
                        sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                        sendData.ExceptionCode = 1;
                        SendException(sendData.ErrorCode, sendData.ExceptionCode, receiveData, sendData, stream, portIn,
                            ipAddressIn);
                    }

                    break;
                //读取输入线圈
                case 2:
                    if (!FunctionCode2Disabled)
                        ReadDiscreteInputs(receiveData, sendData, stream, portIn, ipAddressIn);
                    else
                    {
                        sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                        sendData.ExceptionCode = 1;
                        SendException(sendData.ErrorCode, sendData.ExceptionCode, receiveData, sendData, stream, portIn,
                            ipAddressIn);
                    }

                    break;
                //读取输出寄存器
                case 3:
                    if (!FunctionCode3Disabled)
                        ReadHoldingRegisters(receiveData, sendData, stream, portIn, ipAddressIn);
                    else
                    {
                        sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                        sendData.ExceptionCode = 1;
                        SendException(sendData.ErrorCode, sendData.ExceptionCode, receiveData, sendData, stream, portIn,
                            ipAddressIn);
                    }

                    break;
                // 读取输入寄存器
                case 4:
                    if (!FunctionCode4Disabled)
                        ReadInputRegisters(receiveData, sendData, stream, portIn, ipAddressIn);
                    else
                    {
                        sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                        sendData.ExceptionCode = 1;
                        SendException(sendData.ErrorCode, sendData.ExceptionCode, receiveData, sendData, stream, portIn,
                            ipAddressIn);
                    }

                    break;
                // 写入单个输出线圈
                case 5:
                    if (!FunctionCode5Disabled)
                        WriteSingleCoil(receiveData, sendData, stream, portIn, ipAddressIn);
                    else
                    {
                        sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                        sendData.ExceptionCode = 1;
                        SendException(sendData.ErrorCode, sendData.ExceptionCode, receiveData, sendData, stream, portIn,
                            ipAddressIn);
                    }

                    break;
                // 写入单个输出寄存器
                case 6:
                    if (!FunctionCode6Disabled)
                        WriteSingleRegister(receiveData, sendData, stream, portIn, ipAddressIn);
                    else
                    {
                        sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                        sendData.ExceptionCode = 1;
                        SendException(sendData.ErrorCode, sendData.ExceptionCode, receiveData, sendData, stream, portIn,
                            ipAddressIn);
                    }

                    break;
                // 写入多个输出线圈
                case 15:
                    if (!FunctionCode15Disabled)
                        WriteMultipleCoils(receiveData, sendData, stream, portIn, ipAddressIn);
                    else
                    {
                        sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                        sendData.ExceptionCode = 1;
                        SendException(sendData.ErrorCode, sendData.ExceptionCode, receiveData, sendData, stream, portIn,
                            ipAddressIn);
                    }

                    break;
                // 写入多个输出寄存器
                case 16:
                    if (!FunctionCode16Disabled)
                        WriteMultipleRegisters(receiveData, sendData, stream, portIn, ipAddressIn);
                    else
                    {
                        sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                        sendData.ExceptionCode = 1;
                        SendException(sendData.ErrorCode, sendData.ExceptionCode, receiveData, sendData, stream, portIn,
                            ipAddressIn);
                    }

                    break;
                // 读写多个输出寄存器
                case 23:
                    if (!FunctionCode23Disabled)
                        ReadWriteMultipleRegisters(receiveData, sendData, stream, portIn, ipAddressIn);
                    else
                    {
                        sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                        sendData.ExceptionCode = 1;
                        SendException(sendData.ErrorCode, sendData.ExceptionCode, receiveData, sendData, stream, portIn,
                            ipAddressIn);
                    }

                    break;
                // 未定义的功能码
                default:
                    sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                    sendData.ExceptionCode = 1;
                    SendException(sendData.ErrorCode, sendData.ExceptionCode, receiveData, sendData, stream, portIn,
                        ipAddressIn);
                    break;
            }

            sendData.TimeStamp = DateTime.Now;
        }

        #endregion

        /// <summary>
        /// 读取输出线圈状态
        /// </summary>
        /// <param name="receiveData">接收数据</param>
        /// <param name="sendData">发送数据</param>
        /// <param name="stream">网络流</param>
        /// <param name="portIn">客户端端口</param>
        /// <param name="ipAddressIn">客户端IP</param>
        /// <exception cref="SerialPortNotOpenedException">通讯端口未打开</exception>
        private void ReadCoils(ModbusProtocol receiveData, ModbusProtocol sendData, NetworkStream stream, int portIn,
            IPAddress? ipAddressIn)
        {
            sendData.Response = true;

            sendData.TransactionIdentifier = receiveData.TransactionIdentifier;
            sendData.ProtocolIdentifier = receiveData.ProtocolIdentifier;

            sendData.UnitIdentifier = UnitIdentifier;
            sendData.FunctionCode = receiveData.FunctionCode;
            //非法数据值，由于主站设备试图写入一个超出从站设备可接受范围的值。
            if ((receiveData.Quantity < 1) | (receiveData.Quantity > 0x07D0))
            {
                sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                sendData.ExceptionCode = 3;
            }

            //非法数据地址，主站设备发起了一个超出从站设备寄存器范围的读/写操作。
            if ((receiveData.StartingAddress + 1 + receiveData.Quantity > 65535))
            {
                sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                sendData.ExceptionCode = 2;
            }

            if (sendData.ExceptionCode == 0)
            {
                if ((receiveData.Quantity % 8) == 0)
                    sendData.ByteCount = (byte)(receiveData.Quantity / 8);
                else
                    sendData.ByteCount = (byte)(receiveData.Quantity / 8 + 1);

                sendData.SendCoilValues = new bool[receiveData.Quantity];
                lock (_lockCoils)
                    Array.Copy(CoilsInfo.LocalArray, receiveData.StartingAddress + 1, sendData.SendCoilValues, 0,
                        receiveData.Quantity);
            }

            var data = sendData.ExceptionCode > 0
                ? new byte[9 + 2 * Convert.ToInt32(SerialFlag)]
                : new byte[9 + sendData.ByteCount + 2 * Convert.ToInt32(SerialFlag)];

            sendData.Length = (byte)(data.Length - 6);

            //事务标识符
            var byteData = BitConverter.GetBytes((int)sendData.TransactionIdentifier);
            data[0] = byteData[1];
            data[1] = byteData[0];

            //协议标识符
            byteData = BitConverter.GetBytes((int)sendData.ProtocolIdentifier);
            data[2] = byteData[1];
            data[3] = byteData[0];

            //长度
            byteData = BitConverter.GetBytes((int)sendData.Length);
            data[4] = byteData[1];
            data[5] = byteData[0];

            //单元标识符
            data[6] = sendData.UnitIdentifier;

            //功能码
            data[7] = sendData.FunctionCode;

            //字节数量
            data[8] = sendData.ByteCount;

            if (sendData.ExceptionCode > 0)
            {
                data[7] = sendData.ErrorCode;
                data[8] = sendData.ExceptionCode;
                sendData.SendCoilValues = null;
            }

            if (sendData.SendCoilValues != null)
            {
                for (var i = 0; i < (sendData.ByteCount); i++)
                {
                    byteData = new byte[2];
                    for (var j = 0; j < 8; j++)
                    {
                        var boolValue = sendData.SendCoilValues[i * 8 + j] ? (byte)1 : (byte)0;
                        byteData[1] = (byte)((byteData[1]) | (boolValue << j));
                        if (i * 8 + j + 1 >= sendData.SendCoilValues.Length)
                            break;
                    }

                    data[9 + i] = byteData[1];
                }
            }

            try
            {
                if (SerialFlag)
                {
                    if (_serialPort is { IsOpen: false })
                        throw new SerialPortNotOpenedException("串行端口未打开");
                    //CRC
                    sendData.Crc = ModbusClient.CalculateCrc(data, Convert.ToUInt16(data.Length - 8), 6);
                    byteData = BitConverter.GetBytes((int)sendData.Crc);
                    data[^2] = byteData[0];
                    data[^1] = byteData[1];
                    _serialPort?.Write(data, 6, data.Length - 6);
                }
                else if (UdpFlag)
                {
                    if (ipAddressIn == null)
                        return;
                    var endPoint = new IPEndPoint(ipAddressIn, portIn);
                    _udpClient?.Send(data, data.Length, endPoint);
                }
                else
                {
                    stream.Write(data, 0, data.Length);
                }
            }
            catch (Exception)
            {
                // ignored
            }
        }

        /// <summary>
        /// 读取输入线圈状态
        /// </summary>
        /// <param name="receiveData">接收数据</param>
        /// <param name="sendData">发送数据</param>
        /// <param name="stream">网络流</param>
        /// <param name="portIn">客户端端口</param>
        /// <param name="ipAddressIn">客户端IP</param>
        /// <exception cref="SerialPortNotOpenedException">通讯端口未打开</exception>
        private void ReadDiscreteInputs(ModbusProtocol receiveData, ModbusProtocol sendData, NetworkStream stream,
            int portIn, IPAddress? ipAddressIn)
        {
            sendData.Response = true;

            sendData.TransactionIdentifier = receiveData.TransactionIdentifier;
            sendData.ProtocolIdentifier = receiveData.ProtocolIdentifier;

            sendData.UnitIdentifier = UnitIdentifier;
            sendData.FunctionCode = receiveData.FunctionCode;
            //非法数据值，由于主站设备试图写入一个超出从站设备可接受范围的值。
            if ((receiveData.Quantity < 1) | (receiveData.Quantity > 0x07D0))
            {
                sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                sendData.ExceptionCode = 3;
            }

            //非法数据地址，主站设备发起了一个超出从站设备寄存器范围的读/写操作。
            if (receiveData.StartingAddress + 1 + receiveData.Quantity > 65535)
            {
                sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                sendData.ExceptionCode = 2;
            }

            if (sendData.ExceptionCode == 0)
            {
                if ((receiveData.Quantity % 8) == 0)
                    sendData.ByteCount = (byte)(receiveData.Quantity / 8);
                else
                    sendData.ByteCount = (byte)(receiveData.Quantity / 8 + 1);

                sendData.SendCoilValues = new bool[receiveData.Quantity];
                Array.Copy(DiscreteInputsInfo.LocalArray, receiveData.StartingAddress + 1, sendData.SendCoilValues, 0,
                    receiveData.Quantity);
            }

            var data = sendData.ExceptionCode > 0
                ? new byte[9 + 2 * Convert.ToInt32(SerialFlag)]
                : new byte[9 + sendData.ByteCount + 2 * Convert.ToInt32(SerialFlag)];
            sendData.Length = (byte)(data.Length - 6);

            //事务标识符
            var byteData = BitConverter.GetBytes((int)sendData.TransactionIdentifier);
            data[0] = byteData[1];
            data[1] = byteData[0];

            //协议标识符
            byteData = BitConverter.GetBytes((int)sendData.ProtocolIdentifier);
            data[2] = byteData[1];
            data[3] = byteData[0];

            //长度
            byteData = BitConverter.GetBytes((int)sendData.Length);
            data[4] = byteData[1];
            data[5] = byteData[0];

            //单元标识符
            data[6] = sendData.UnitIdentifier;

            //功能码
            data[7] = sendData.FunctionCode;

            //字节数量
            data[8] = sendData.ByteCount;

            if (sendData.ExceptionCode > 0)
            {
                data[7] = sendData.ErrorCode;
                data[8] = sendData.ExceptionCode;
                sendData.SendCoilValues = null;
            }

            if (sendData.SendCoilValues != null)
                for (var i = 0; i < (sendData.ByteCount); i++)
                {
                    byteData = new byte[2];
                    for (var j = 0; j < 8; j++)
                    {
                        var boolValue = sendData.SendCoilValues[i * 8 + j] ? (byte)1 : (byte)0;
                        byteData[1] = (byte)((byteData[1]) | (boolValue << j));
                        if (i * 8 + j + 1 >= sendData.SendCoilValues.Length)
                            break;
                    }

                    data[9 + i] = byteData[1];
                }

            try
            {
                if (SerialFlag)
                {
                    if (_serialPort is { IsOpen: false })
                        throw new SerialPortNotOpenedException("串行端口未打开");
                    //CRC
                    sendData.Crc = ModbusClient.CalculateCrc(data, Convert.ToUInt16(data.Length - 8), 6);
                    byteData = BitConverter.GetBytes((int)sendData.Crc);
                    data[^2] = byteData[0];
                    data[^1] = byteData[1];
                    _serialPort?.Write(data, 6, data.Length - 6);
                }
                else if (UdpFlag)
                {
                    if (ipAddressIn == null)
                        return;
                    var endPoint = new IPEndPoint(ipAddressIn, portIn);
                    _udpClient?.Send(data, data.Length, endPoint);
                }
                else
                {
                    stream.Write(data, 0, data.Length);
                }
            }
            catch (Exception)
            {
                // ignored
            }
        }

        /// <summary>
        /// 读取输出寄存器
        /// </summary>
        /// <param name="receiveData">接收数据</param>
        /// <param name="sendData">发送数据</param>
        /// <param name="stream">网络流</param>
        /// <param name="portIn">客户端端口</param>
        /// <param name="ipAddressIn">客户端IP</param>
        /// <exception cref="SerialPortNotOpenedException">通讯端口未打开</exception>
        private void ReadHoldingRegisters(ModbusProtocol receiveData, ModbusProtocol sendData, NetworkStream stream,
            int portIn, IPAddress? ipAddressIn)
        {
            sendData.Response = true;

            sendData.TransactionIdentifier = receiveData.TransactionIdentifier;
            sendData.ProtocolIdentifier = receiveData.ProtocolIdentifier;

            sendData.UnitIdentifier = UnitIdentifier;
            sendData.FunctionCode = receiveData.FunctionCode;

            //非法数据值，由于主站设备试图写入一个超出从站设备可接受范围的值。
            if ((receiveData.Quantity < 1) | (receiveData.Quantity > 0x007D))
            {
                sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                sendData.ExceptionCode = 3;
            }

            //非法数据地址，主站设备发起了一个超出从站设备寄存器范围的读/写操作。
            if (receiveData.StartingAddress + 1 + receiveData.Quantity > 65535)
            {
                sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                sendData.ExceptionCode = 2;
            }

            if (sendData.ExceptionCode == 0)
            {
                sendData.ByteCount = (byte)(2 * receiveData.Quantity);
                sendData.SendRegisterValues = new short[receiveData.Quantity];
                lock (_lockHoldingRegisters)
                    Buffer.BlockCopy(HoldingRegistersInfo.LocalArray, receiveData.StartingAddress * 2 + 2,
                        sendData.SendRegisterValues, 0, receiveData.Quantity * 2);
            }

            if (sendData.ExceptionCode > 0)
                sendData.Length = 0x03;
            else
                sendData.Length = (ushort)(0x03 + sendData.ByteCount);

            var data = sendData.ExceptionCode > 0
                ? new byte[9 + 2 * Convert.ToInt32(SerialFlag)]
                : new byte[9 + sendData.ByteCount + 2 * Convert.ToInt32(SerialFlag)];
            sendData.Length = (byte)(data.Length - 6);

            //事务标识符
            var byteData = BitConverter.GetBytes((int)sendData.TransactionIdentifier);
            data[0] = byteData[1];
            data[1] = byteData[0];

            //协议标识符
            byteData = BitConverter.GetBytes((int)sendData.ProtocolIdentifier);
            data[2] = byteData[1];
            data[3] = byteData[0];

            //长度
            byteData = BitConverter.GetBytes((int)sendData.Length);
            data[4] = byteData[1];
            data[5] = byteData[0];

            //单元标识符
            data[6] = sendData.UnitIdentifier;

            //功能码
            data[7] = sendData.FunctionCode;

            //字节数量
            data[8] = sendData.ByteCount;

            if (sendData.ExceptionCode > 0)
            {
                data[7] = sendData.ErrorCode;
                data[8] = sendData.ExceptionCode;
                sendData.SendRegisterValues = null;
            }


            if (sendData.SendRegisterValues != null)
            {
                for (var i = 0; i < (sendData.ByteCount / 2); i++)
                {
                    byteData = BitConverter.GetBytes(sendData.SendRegisterValues[i]);
                    data[9 + i * 2] = byteData[1];
                    data[10 + i * 2] = byteData[0];
                }
            }

            try
            {
                if (SerialFlag)
                {
                    if (_serialPort is { IsOpen: false })
                        throw new SerialPortNotOpenedException("串行端口未打开");
                    //CRC 校验
                    sendData.Crc = ModbusClient.CalculateCrc(data, Convert.ToUInt16(data.Length - 8), 6);
                    byteData = BitConverter.GetBytes((int)sendData.Crc);
                    data[^2] = byteData[0];
                    data[^1] = byteData[1];
                    _serialPort?.Write(data, 6, data.Length - 6);
                }
                else if (UdpFlag)
                {
                    if (ipAddressIn == null)
                        return;
                    var endPoint = new IPEndPoint(ipAddressIn, portIn);
                    _udpClient?.Send(data, data.Length, endPoint);
                }
                else
                {
                    stream.Write(data, 0, data.Length);
                }
            }
            catch (Exception)
            {
                // ignored
            }
        }

        /// <summary>
        /// 读取输入寄存器
        /// </summary>
        /// <param name="receiveData">接收数据</param>
        /// <param name="sendData">发送数据</param>
        /// <param name="stream">网络流</param>
        /// <param name="portIn">客户端端口</param>
        /// <param name="ipAddressIn">客户端IP</param>
        /// <exception cref="SerialPortNotOpenedException">通讯端口未打开</exception>
        private void ReadInputRegisters(ModbusProtocol receiveData, ModbusProtocol sendData, NetworkStream stream,
            int portIn, IPAddress? ipAddressIn)
        {
            sendData.Response = true;

            sendData.TransactionIdentifier = receiveData.TransactionIdentifier;
            sendData.ProtocolIdentifier = receiveData.ProtocolIdentifier;

            sendData.UnitIdentifier = UnitIdentifier;
            sendData.FunctionCode = receiveData.FunctionCode;
            //非法数据值，由于主站设备试图写入一个超出从站设备可接受范围的值。
            if ((receiveData.Quantity < 1) | (receiveData.Quantity > 0x007D))
            {
                sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                sendData.ExceptionCode = 3;
            }

            //非法数据地址，主站设备发起了一个超出从站设备寄存器范围的读/写操作。
            if (receiveData.StartingAddress + 1 + receiveData.Quantity > 65535)
            {
                sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                sendData.ExceptionCode = 2;
            }

            if (sendData.ExceptionCode == 0)
            {
                sendData.ByteCount = (byte)(2 * receiveData.Quantity);
                sendData.SendRegisterValues = new short[receiveData.Quantity];
                Buffer.BlockCopy(InputRegistersInfo.LocalArray, receiveData.StartingAddress * 2 + 2,
                    sendData.SendRegisterValues, 0, receiveData.Quantity * 2);
            }

            if (sendData.ExceptionCode > 0)
                sendData.Length = 0x03;
            else
                sendData.Length = (ushort)(0x03 + sendData.ByteCount);

            var data = sendData.ExceptionCode > 0
                ? new byte[9 + 2 * Convert.ToInt32(SerialFlag)]
                : new byte[9 + sendData.ByteCount + 2 * Convert.ToInt32(SerialFlag)];
            sendData.Length = (byte)(data.Length - 6);

            //事务标识符
            var byteData = BitConverter.GetBytes((int)sendData.TransactionIdentifier);
            data[0] = byteData[1];
            data[1] = byteData[0];

            //协议标识符
            byteData = BitConverter.GetBytes((int)sendData.ProtocolIdentifier);
            data[2] = byteData[1];
            data[3] = byteData[0];

            //长度
            byteData = BitConverter.GetBytes((int)sendData.Length);
            data[4] = byteData[1];
            data[5] = byteData[0];

            //单元标识符
            data[6] = sendData.UnitIdentifier;

            //功能码
            data[7] = sendData.FunctionCode;

            //字节数量
            data[8] = sendData.ByteCount;


            if (sendData.ExceptionCode > 0)
            {
                data[7] = sendData.ErrorCode;
                data[8] = sendData.ExceptionCode;
                sendData.SendRegisterValues = null;
            }


            if (sendData.SendRegisterValues != null)
            {
                for (var i = 0; i < (sendData.ByteCount / 2); i++)
                {
                    byteData = BitConverter.GetBytes(sendData.SendRegisterValues[i]);
                    data[9 + i * 2] = byteData[1];
                    data[10 + i * 2] = byteData[0];
                }
            }

            try
            {
                if (SerialFlag)
                {
                    if (_serialPort is { IsOpen: false })
                        throw new SerialPortNotOpenedException("串行端口未打开");
                    //CRC 校验
                    sendData.Crc = ModbusClient.CalculateCrc(data, Convert.ToUInt16(data.Length - 8), 6);
                    byteData = BitConverter.GetBytes((int)sendData.Crc);
                    data[^2] = byteData[0];
                    data[^1] = byteData[1];
                    _serialPort?.Write(data, 6, data.Length - 6);
                }
                else if (UdpFlag)
                {
                    if (ipAddressIn == null)
                        return;
                    var endPoint = new IPEndPoint(ipAddressIn, portIn);
                    _udpClient?.Send(data, data.Length, endPoint);
                }
                else
                {
                    stream.Write(data, 0, data.Length);
                }
            }
            catch (Exception)
            {
                // ignored
            }
        }

        /// <summary>
        /// 写入单个输出线圈
        /// </summary>
        /// <param name="receiveData">接收数据</param>
        /// <param name="sendData">发送数据</param>
        /// <param name="stream">网络流</param>
        /// <param name="portIn">客户端端口</param>
        /// <param name="ipAddressIn">客户端IP</param>
        /// <exception cref="SerialPortNotOpenedException">通讯端口未打开</exception>
        private void WriteSingleCoil(ModbusProtocol receiveData, ModbusProtocol sendData, NetworkStream stream,
            int portIn, IPAddress? ipAddressIn)
        {
            sendData.Response = true;

            sendData.TransactionIdentifier = receiveData.TransactionIdentifier;
            sendData.ProtocolIdentifier = receiveData.ProtocolIdentifier;

            sendData.UnitIdentifier = UnitIdentifier;
            sendData.FunctionCode = receiveData.FunctionCode;
            sendData.StartingAddress = receiveData.StartingAddress;
            sendData.ReceiveCoilValues = receiveData.ReceiveCoilValues;
#pragma warning disable CS8602 // 解引用可能出现空引用。
            if ((receiveData.ReceiveCoilValues[0] != 0x0000)
                & (receiveData.ReceiveCoilValues[0] != 0xFF00))
            {
                sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                sendData.ExceptionCode = 3;
            }
#pragma warning restore CS8602 // 解引用可能出现空引用。
            if ((receiveData.StartingAddress + 1) > 65535)
            {
                sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                sendData.ExceptionCode = 2;
            }

            if (sendData.ExceptionCode == 0)
            {
                if (receiveData.ReceiveCoilValues[0] == 0xFF00)
                {
                    lock (_lockCoils)
                        CoilsInfo[receiveData.StartingAddress + 1] = true;
                }

                if (receiveData.ReceiveCoilValues[0] == 0x0000)
                {
                    lock (_lockCoils)
                        CoilsInfo[receiveData.StartingAddress + 1] = false;
                }
            }

            sendData.Length = sendData.ExceptionCode > 0 ? (ushort)0x03 : (ushort)0x06;
            var data = sendData.ExceptionCode > 0
                ? new byte[9 + 2 * Convert.ToInt32(SerialFlag)]
                : new byte[12 + 2 * Convert.ToInt32(SerialFlag)];
            sendData.Length = (byte)(data.Length - 6);

            //事务标识符
            var byteData = BitConverter.GetBytes((int)sendData.TransactionIdentifier);
            data[0] = byteData[1];
            data[1] = byteData[0];

            //协议标识符
            byteData = BitConverter.GetBytes((int)sendData.ProtocolIdentifier);
            data[2] = byteData[1];
            data[3] = byteData[0];

            //长度
            byteData = BitConverter.GetBytes((int)sendData.Length);
            data[4] = byteData[1];
            data[5] = byteData[0];

            //单元标识符
            data[6] = sendData.UnitIdentifier;

            //功能码
            data[7] = sendData.FunctionCode;

            if (sendData.ExceptionCode > 0)
            {
                data[7] = sendData.ErrorCode;
                data[8] = sendData.ExceptionCode;
                sendData.SendRegisterValues = null;
            }
            else
            {
                byteData = BitConverter.GetBytes((int)receiveData.StartingAddress);
                data[8] = byteData[1];
                data[9] = byteData[0];
                byteData = BitConverter.GetBytes((int)receiveData.ReceiveCoilValues[0]);
                data[10] = byteData[1];
                data[11] = byteData[0];
            }


            try
            {
                if (SerialFlag)
                {
                    if (_serialPort is { IsOpen: false })
                        throw new SerialPortNotOpenedException("串行端口未打开");
                    //CRC 校验
                    sendData.Crc = ModbusClient.CalculateCrc(data, Convert.ToUInt16(data.Length - 8), 6);
                    byteData = BitConverter.GetBytes((int)sendData.Crc);
                    data[^2] = byteData[0];
                    data[^1] = byteData[1];
                    _serialPort?.Write(data, 6, data.Length - 6);
                }
                else if (UdpFlag)
                {
                    if (ipAddressIn != null)
                    {
                        var endPoint = new IPEndPoint(ipAddressIn, portIn);
                        _udpClient?.Send(data, data.Length, endPoint);
                    }
                }
                else
                {
                    stream.Write(data, 0, data.Length);
                }
            }
            catch (Exception)
            {
                // ignored
            }

            CoilsChanged?.Invoke(receiveData.StartingAddress + 1, 1);
        }

        /// <summary>
        /// 写入单个输出寄存器
        /// </summary>
        /// <param name="receiveData">接收数据</param>
        /// <param name="sendData">发送数据</param>
        /// <param name="stream">网络流</param>
        /// <param name="portIn">客户端端口</param>
        /// <param name="ipAddressIn">客户端IP</param>
        /// <exception cref="SerialPortNotOpenedException">通讯端口未打开</exception>
        private void WriteSingleRegister(ModbusProtocol receiveData, ModbusProtocol sendData, NetworkStream stream,
            int portIn, IPAddress? ipAddressIn)
        {
            sendData.Response = true;

            sendData.TransactionIdentifier = receiveData.TransactionIdentifier;
            sendData.ProtocolIdentifier = receiveData.ProtocolIdentifier;

            sendData.UnitIdentifier = UnitIdentifier;
            sendData.FunctionCode = receiveData.FunctionCode;
            sendData.StartingAddress = receiveData.StartingAddress;
            sendData.ReceiveRegisterValues = receiveData.ReceiveRegisterValues;


            // if ((receiveData.ReceiveRegisterValues[0] < 0x0000) |
            //     (receiveData.ReceiveRegisterValues[0] > 0xFFFF))
            // {
            //     sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
            //     sendData.ExceptionCode = 3;
            // }

            if (receiveData.StartingAddress + 1 > 65535)
            {
                sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                sendData.ExceptionCode = 2;
            }

            if (sendData.ExceptionCode == 0)
            {
                lock (_lockHoldingRegisters)
                    if (receiveData.ReceiveRegisterValues != null)
                        HoldingRegistersInfo[receiveData.StartingAddress + 1] =
                            unchecked((short)receiveData.ReceiveRegisterValues[0]);
            }

            sendData.Length = sendData.ExceptionCode > 0 ? (ushort)0x03 : (ushort)0x06;

            var data = sendData.ExceptionCode > 0
                ? new byte[9 + 2 * Convert.ToInt32(SerialFlag)]
                : new byte[12 + 2 * Convert.ToInt32(SerialFlag)];

            sendData.Length = (byte)(data.Length - 6);


            //事务标识符
            var byteData = BitConverter.GetBytes((int)sendData.TransactionIdentifier);
            data[0] = byteData[1];
            data[1] = byteData[0];

            //协议标识符
            byteData = BitConverter.GetBytes((int)sendData.ProtocolIdentifier);
            data[2] = byteData[1];
            data[3] = byteData[0];

            //长度
            byteData = BitConverter.GetBytes((int)sendData.Length);
            data[4] = byteData[1];
            data[5] = byteData[0];

            //单元标识符
            data[6] = sendData.UnitIdentifier;

            //功能码
            data[7] = sendData.FunctionCode;


            if (sendData.ExceptionCode > 0)
            {
                data[7] = sendData.ErrorCode;
                data[8] = sendData.ExceptionCode;
                sendData.SendRegisterValues = null;
            }
            else
            {
                byteData = BitConverter.GetBytes((int)receiveData.StartingAddress);
                data[8] = byteData[1];
                data[9] = byteData[0];
                if (receiveData.ReceiveRegisterValues != null)
                    byteData = BitConverter.GetBytes((int)receiveData.ReceiveRegisterValues[0]);
                data[10] = byteData[1];
                data[11] = byteData[0];
            }


            try
            {
                if (SerialFlag)
                {
                    if (_serialPort is { IsOpen: false })
                        throw new SerialPortNotOpenedException("串行端口未打开");
                    //CRC 校验
                    sendData.Crc = ModbusClient.CalculateCrc(data, Convert.ToUInt16(data.Length - 8), 6);
                    byteData = BitConverter.GetBytes((int)sendData.Crc);
                    data[^2] = byteData[0];
                    data[^1] = byteData[1];
                    _serialPort?.Write(data, 6, data.Length - 6);
                }
                else if (UdpFlag)
                {
                    if (ipAddressIn != null)
                    {
                        var endPoint = new IPEndPoint(ipAddressIn, portIn);
                        _udpClient?.Send(data, data.Length, endPoint);
                    }
                }
                else
                {
                    stream.Write(data, 0, data.Length);
                }
            }
            catch (Exception)
            {
                // ignored
            }

            HoldingRegistersChanged?.Invoke(receiveData.StartingAddress + 1, 1);
        }

        /// <summary>
        /// 写入多个输出线圈
        /// </summary>
        /// <param name="receiveData">接收数据</param>
        /// <param name="sendData">发送数据</param>
        /// <param name="stream">网络流</param>
        /// <param name="portIn">客户端端口</param>
        /// <param name="ipAddressIn">客户端IP</param>
        /// <exception cref="SerialPortNotOpenedException">通讯端口未打开</exception>
        private void WriteMultipleCoils(ModbusProtocol receiveData, ModbusProtocol sendData, NetworkStream stream,
            int portIn, IPAddress? ipAddressIn)
        {
            sendData.Response = true;

            sendData.TransactionIdentifier = receiveData.TransactionIdentifier;
            sendData.ProtocolIdentifier = receiveData.ProtocolIdentifier;

            sendData.UnitIdentifier = UnitIdentifier;
            sendData.FunctionCode = receiveData.FunctionCode;
            sendData.StartingAddress = receiveData.StartingAddress;
            sendData.Quantity = receiveData.Quantity;

            if ((receiveData.Quantity == 0x0000) | (receiveData.Quantity > 0x07B0))
            {
                sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                sendData.ExceptionCode = 3;
            }

            if (receiveData.StartingAddress + 1 + receiveData.Quantity > 65535)
            {
                sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                sendData.ExceptionCode = 2;
            }

            if (sendData.ExceptionCode == 0)
            {
                lock (_lockCoils)
                {
                    for (var i = 0; i < receiveData.Quantity; i++)
                    {
                        var shift = i % 16;
                        var mask = 0x1;
                        mask = mask << (shift);
                        if (receiveData.ReceiveCoilValues != null &&
                            (receiveData.ReceiveCoilValues[i / 16] & (ushort)mask) == 0)
                            CoilsInfo[receiveData.StartingAddress + i + 1] = false;
                        else

                            CoilsInfo[receiveData.StartingAddress + i + 1] = true;
                    }
                }
            }

            sendData.Length = sendData.ExceptionCode > 0 ? (ushort)0x03 : (ushort)0x06;
            var data = sendData.ExceptionCode > 0
                ? new byte[9 + 2 * Convert.ToInt32(SerialFlag)]
                : new byte[12 + 2 * Convert.ToInt32(SerialFlag)];
            sendData.Length = (byte)(data.Length - 6);

            //事务标识符
            var byteData = BitConverter.GetBytes((int)sendData.TransactionIdentifier);
            data[0] = byteData[1];
            data[1] = byteData[0];

            //协议标识符
            byteData = BitConverter.GetBytes((int)sendData.ProtocolIdentifier);
            data[2] = byteData[1];
            data[3] = byteData[0];

            //长度
            byteData = BitConverter.GetBytes((int)sendData.Length);
            data[4] = byteData[1];
            data[5] = byteData[0];

            //单元标识符
            data[6] = sendData.UnitIdentifier;

            //功能码
            data[7] = sendData.FunctionCode;


            if (sendData.ExceptionCode > 0)
            {
                data[7] = sendData.ErrorCode;
                data[8] = sendData.ExceptionCode;
                sendData.SendRegisterValues = null;
            }
            else
            {
                byteData = BitConverter.GetBytes((int)receiveData.StartingAddress);
                data[8] = byteData[1];
                data[9] = byteData[0];
                byteData = BitConverter.GetBytes((int)receiveData.Quantity);
                data[10] = byteData[1];
                data[11] = byteData[0];
            }


            try
            {
                if (SerialFlag)
                {
                    if (_serialPort is { IsOpen: false })
                        throw new SerialPortNotOpenedException("串行端口未打开");
                    //CRC 校验
                    sendData.Crc = ModbusClient.CalculateCrc(data, Convert.ToUInt16(data.Length - 8), 6);
                    byteData = BitConverter.GetBytes((int)sendData.Crc);
                    data[^2] = byteData[0];
                    data[^1] = byteData[1];
                    _serialPort?.Write(data, 6, data.Length - 6);
                }
                else if (UdpFlag)
                {
                    if (ipAddressIn != null)
                    {
                        var endPoint = new IPEndPoint(ipAddressIn, portIn);
                        _udpClient?.Send(data, data.Length, endPoint);
                    }
                }
                else
                {
                    stream.Write(data, 0, data.Length);
                }
            }
            catch (Exception)
            {
                // ignored
            }

            CoilsChanged?.Invoke(receiveData.StartingAddress + 1, receiveData.Quantity);
        }

        /// <summary>
        /// 写入多个输出寄存器
        /// </summary>
        /// <param name="receiveData">接收数据</param>
        /// <param name="sendData">发送数据</param>
        /// <param name="stream">网络流</param>
        /// <param name="portIn">客户端端口</param>
        /// <param name="ipAddressIn">客户端IP</param>
        /// <exception cref="SerialPortNotOpenedException">通讯端口未打开</exception>
        private void WriteMultipleRegisters(ModbusProtocol receiveData, ModbusProtocol sendData, NetworkStream stream,
            int portIn, IPAddress? ipAddressIn)
        {
            sendData.Response = true;

            sendData.TransactionIdentifier = receiveData.TransactionIdentifier;
            sendData.ProtocolIdentifier = receiveData.ProtocolIdentifier;

            sendData.UnitIdentifier = UnitIdentifier;
            sendData.FunctionCode = receiveData.FunctionCode;
            sendData.StartingAddress = receiveData.StartingAddress;
            sendData.Quantity = receiveData.Quantity;

            if ((receiveData.Quantity == 0x0000) | (receiveData.Quantity > 0x07B0))
            {
                sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                sendData.ExceptionCode = 3;
            }

            if (receiveData.StartingAddress + 1 + receiveData.Quantity > 65535)
            {
                sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                sendData.ExceptionCode = 2;
            }

            if (sendData.ExceptionCode == 0)
            {
                lock (_lockHoldingRegisters)
                {
                    for (var i = 0; i < receiveData.Quantity; i++)
                    {
                        if (receiveData.ReceiveRegisterValues != null)
                        {
                            HoldingRegistersInfo[receiveData.StartingAddress + i + 1] =
                                unchecked((short)receiveData.ReceiveRegisterValues[i]);
                        }
                    }
                }
            }

            sendData.Length = sendData.ExceptionCode > 0 ? (ushort)0x03 : (ushort)0x06;
            {
                var data = sendData.ExceptionCode > 0
                    ? new byte[9 + 2 * Convert.ToInt32(SerialFlag)]
                    : new byte[12 + 2 * Convert.ToInt32(SerialFlag)];

                sendData.Length = (byte)(data.Length - 6);

                //事务标识符
                var byteData = BitConverter.GetBytes((int)sendData.TransactionIdentifier);
                data[0] = byteData[1];
                data[1] = byteData[0];

                //协议标识符
                byteData = BitConverter.GetBytes((int)sendData.ProtocolIdentifier);
                data[2] = byteData[1];
                data[3] = byteData[0];

                //长度
                byteData = BitConverter.GetBytes((int)sendData.Length);
                data[4] = byteData[1];
                data[5] = byteData[0];

                //单元标识符
                data[6] = sendData.UnitIdentifier;

                //功能码
                data[7] = sendData.FunctionCode;


                if (sendData.ExceptionCode > 0)
                {
                    data[7] = sendData.ErrorCode;
                    data[8] = sendData.ExceptionCode;
                    sendData.SendRegisterValues = null;
                }
                else
                {
                    byteData = BitConverter.GetBytes((int)receiveData.StartingAddress);
                    data[8] = byteData[1];
                    data[9] = byteData[0];
                    byteData = BitConverter.GetBytes((int)receiveData.Quantity);
                    data[10] = byteData[1];
                    data[11] = byteData[0];
                }


                try
                {
                    if (SerialFlag)
                    {
                        if (_serialPort is { IsOpen: false })
                            throw new SerialPortNotOpenedException("串行端口未打开");
                        //CRC 校验
                        sendData.Crc = ModbusClient.CalculateCrc(data, Convert.ToUInt16(data.Length - 8), 6);
                        byteData = BitConverter.GetBytes((int)sendData.Crc);
                        data[^2] = byteData[0];
                        data[^1] = byteData[1];
                        _serialPort?.Write(data, 6, data.Length - 6);
                    }
                    else if (UdpFlag)
                    {
                        if (ipAddressIn != null)
                        {
                            var endPoint = new IPEndPoint(ipAddressIn, portIn);
                            _udpClient?.Send(data, data.Length, endPoint);
                        }
                    }
                    else
                    {
                        stream.Write(data, 0, data.Length);
                    }
                }
                catch (Exception)
                {
                    // ignored
                }

                HoldingRegistersChanged?.Invoke(receiveData.StartingAddress + 1, receiveData.Quantity);
            }
        }

        /// <summary>
        /// 写入读取多个输出寄存器
        /// </summary>
        /// <param name="receiveData">接收数据</param>
        /// <param name="sendData">发送数据</param>
        /// <param name="stream">网络流</param>
        /// <param name="portIn">客户端端口</param>
        /// <param name="ipAddressIn">客户端IP</param>
        /// <exception cref="SerialPortNotOpenedException">通讯端口未打开</exception>
        private void ReadWriteMultipleRegisters(ModbusProtocol receiveData, ModbusProtocol sendData,
            NetworkStream stream, int portIn, IPAddress? ipAddressIn)
        {
            sendData.Response = true;

            sendData.TransactionIdentifier = receiveData.TransactionIdentifier;
            sendData.ProtocolIdentifier = receiveData.ProtocolIdentifier;

            sendData.UnitIdentifier = UnitIdentifier;
            sendData.FunctionCode = receiveData.FunctionCode;


            if ((receiveData.QuantityRead < 0x0001) | (receiveData.QuantityRead > 0x007D) |
                (receiveData.QuantityWrite < 0x0001) | (receiveData.QuantityWrite > 0x0079) |
                (receiveData.ByteCount != receiveData.QuantityWrite * 2))
            {
                sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                sendData.ExceptionCode = 3;
            }

            if ((receiveData.StartingAddressRead + 1 + receiveData.QuantityRead > 65535) |
                (receiveData.StartingAddressWrite + 1 + receiveData.QuantityWrite > 65535))
            {
                sendData.ErrorCode = (byte)(receiveData.FunctionCode + 0x80);
                sendData.ExceptionCode = 2;
            }

            if (sendData.ExceptionCode == 0)
            {
                sendData.SendRegisterValues = new short[receiveData.QuantityRead];
                lock (_lockHoldingRegisters)
                    Buffer.BlockCopy(HoldingRegistersInfo.LocalArray, receiveData.StartingAddressRead * 2 + 2,
                        sendData.SendRegisterValues, 0, receiveData.QuantityRead * 2);

                lock (HoldingRegistersInfo)
                {
                    for (var i = 0; i < receiveData.QuantityWrite; i++)
                    {
                        if (receiveData.ReceiveRegisterValues != null)
                            HoldingRegistersInfo[receiveData.StartingAddressWrite + i + 1] =
                                unchecked((short)receiveData.ReceiveRegisterValues[i]);
                    }
                }


                sendData.ByteCount = (byte)(2 * receiveData.QuantityRead);
            }

            sendData.Length = sendData.ExceptionCode > 0
                ? (ushort)0x03
                : Convert.ToUInt16(3 + 2 * receiveData.QuantityRead);


            var data = sendData.ExceptionCode > 0
                ? new byte[9 + 2 * Convert.ToInt32(SerialFlag)]
                : new byte[9 + sendData.ByteCount + 2 * Convert.ToInt32(SerialFlag)];

            //事务标识符
            var byteData = BitConverter.GetBytes((int)sendData.TransactionIdentifier);
            data[0] = byteData[1];
            data[1] = byteData[0];

            //协议标识符
            byteData = BitConverter.GetBytes((int)sendData.ProtocolIdentifier);
            data[2] = byteData[1];
            data[3] = byteData[0];

            //长度
            byteData = BitConverter.GetBytes((int)sendData.Length);
            data[4] = byteData[1];
            data[5] = byteData[0];

            //单元标识符
            data[6] = sendData.UnitIdentifier;

            //功能码
            data[7] = sendData.FunctionCode;

            //字节数量
            data[8] = sendData.ByteCount;


            if (sendData.ExceptionCode > 0)
            {
                data[7] = sendData.ErrorCode;
                data[8] = sendData.ExceptionCode;
                sendData.SendRegisterValues = null;
            }
            else
            {
                if (sendData.SendRegisterValues != null)
                    for (var i = 0; i < sendData.ByteCount / 2; i++)
                    {
                        byteData = BitConverter.GetBytes(sendData.SendRegisterValues[i]);
                        data[9 + i * 2] = byteData[1];
                        data[10 + i * 2] = byteData[0];
                    }
            }


            try
            {
                if (SerialFlag)
                {
                    if (_serialPort is { IsOpen: false })
                        throw new SerialPortNotOpenedException("串行端口未打开");
                    //CRC 校验
                    sendData.Crc = ModbusClient.CalculateCrc(data, Convert.ToUInt16(data.Length - 8), 6);
                    byteData = BitConverter.GetBytes((int)sendData.Crc);
                    data[^2] = byteData[0];
                    data[^1] = byteData[1];
                    _serialPort?.Write(data, 6, data.Length - 6);
                }
                else if (UdpFlag)
                {
                    if (ipAddressIn != null)
                    {
                        var endPoint = new IPEndPoint(ipAddressIn, portIn);
                        _udpClient?.Send(data, data.Length, endPoint);
                    }
                }
                else
                {
                    stream.Write(data, 0, data.Length);
                }
            }
            catch (Exception)
            {
                // ignored
            }

            HoldingRegistersChanged?.Invoke(receiveData.StartingAddressWrite + 1, receiveData.QuantityWrite);
        }

        private void SendException(int errorCode, int exceptionCode, ModbusProtocol receiveData,
            ModbusProtocol sendData, NetworkStream stream, int portIn, IPAddress? ipAddressIn)
        {
            sendData.Response = true;

            sendData.TransactionIdentifier = receiveData.TransactionIdentifier;
            sendData.ProtocolIdentifier = receiveData.ProtocolIdentifier;

            sendData.UnitIdentifier = receiveData.UnitIdentifier;
            sendData.ErrorCode = (byte)errorCode;
            sendData.ExceptionCode = (byte)exceptionCode;

            if (sendData.ExceptionCode > 0)
                sendData.Length = 0x03;
            else
                sendData.Length = (ushort)(0x03 + sendData.ByteCount);


            var data = sendData.ExceptionCode > 0
                ? new byte[9 + 2 * Convert.ToInt32(SerialFlag)]
                : new byte[9 + sendData.ByteCount + 2 * Convert.ToInt32(SerialFlag)];
            sendData.Length = (byte)(data.Length - 6);

            //事务标识符
            var byteData = BitConverter.GetBytes((int)sendData.TransactionIdentifier);
            data[0] = byteData[1];
            data[1] = byteData[0];

            //协议标识符
            byteData = BitConverter.GetBytes((int)sendData.ProtocolIdentifier);
            data[2] = byteData[1];
            data[3] = byteData[0];

            //长度
            byteData = BitConverter.GetBytes((int)sendData.Length);
            data[4] = byteData[1];
            data[5] = byteData[0];

            //单元标识符
            data[6] = sendData.UnitIdentifier;

            //错误代码
            data[7] = sendData.ErrorCode;
            data[8] = sendData.ExceptionCode;


            try
            {
                if (SerialFlag)
                {
                    if (_serialPort is { IsOpen: false })
                        throw new SerialPortNotOpenedException("串行端口未打开");
                    //CRC 校验
                    sendData.Crc = ModbusClient.CalculateCrc(data, Convert.ToUInt16(data.Length - 8), 6);
                    byteData = BitConverter.GetBytes((int)sendData.Crc);
                    data[^2] = byteData[0];
                    data[^1] = byteData[1];
                    _serialPort?.Write(data, 6, data.Length - 6);
                }
                else if (UdpFlag)
                {
                    if (ipAddressIn == null)
                        return;
                    var endPoint = new IPEndPoint(ipAddressIn, portIn);
                    _udpClient?.Send(data, data.Length, endPoint);
                }
                else
                {
                    stream.Write(data, 0, data.Length);
                }
            }
            catch (Exception)
            {
                // ignored
            }
        }

        private void CreateLogData(ModbusProtocol receiveData, ModbusProtocol sendData)
        {
            for (var i = 0; i < 98; i++)
            {
                ModbusLogData[99 - i] = ModbusLogData[99 - i - 2];
            }

            ModbusLogData[0] = receiveData;
            ModbusLogData[1] = sendData;
        }

        /// <summary>
        /// 当前连接的数量
        /// </summary>
        public int NumberOfConnections { get; private set; }

        /// <summary>
        /// 最近100条的Modbus协议数据
        /// </summary>
        public ModbusProtocol[] ModbusLogData { get; } = new ModbusProtocol[100];
        
        /// <summary>
        /// 通讯端口号
        /// </summary>
        public int Port { get; set; } = 502;

        /// <summary>
        /// UDP 模式标识
        /// </summary>
        public bool UdpFlag { get; set; }

        /// <summary>
        /// RTU工作模式标识
        /// </summary>
        public bool SerialFlag { get; set; }

        /// <summary>
        /// 波特率
        /// </summary>
        public int BaudRate { get; set; } = 9600;

        /// <summary>
        /// 奇偶校验位
        /// </summary>
        public Parity Parity { get; set; } = Parity.Even;

        /// <summary>
        /// 停止位
        /// </summary>
        public StopBits StopBits { get; set; } = StopBits.One;

        /// <summary>
        /// 串口号
        /// </summary>
        public string? SerialPortName
        {
            get => _serialPortName;
            set
            {
                if (_serialPortName != value)
                {
                    _serialPortName = value;
                    _portChanged = true;
                }
                else
                {
                    _serialPortName = value;
                }

                SerialFlag = _serialPortName != null;
            }
        }

        /// <summary>
        /// 单位标识符
        /// </summary>
        public byte UnitIdentifier { get; set; } = 1;

        public class HoldingRegisters(ModbusServer modbusServer)
        {
            public readonly short[] LocalArray = new short[65535];
            private ModbusServer _modbusServer = modbusServer;

            public short this[int x]
            {
                get => LocalArray[x];
                set => LocalArray[x] = value;
            }
        }

        public class InputRegisters(ModbusServer modbusServer)
        {
            public readonly short[] LocalArray = new short[65535];
            private ModbusServer _modbusServer = modbusServer;

            public short this[int x]
            {
                get => LocalArray[x];
                set => LocalArray[x] = value;
            }
        }

        public class Coils(ModbusServer modbusServer)
        {
            public readonly bool[] LocalArray = new bool[65535];
            private ModbusServer _modbusServer = modbusServer;

            public bool this[int x]
            {
                get => LocalArray[x];
                set => LocalArray[x] = value;
            }
        }

        public class DiscreteInputs(ModbusServer modbusServer)
        {
            public readonly bool[] LocalArray = new bool[65535];
            private ModbusServer _modbusServer = modbusServer;

            public bool this[int x]
            {
                get => LocalArray[x];
                set => LocalArray[x] = value;
            }
        }
    }
}