using System.IO.Ports;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using ModbusTCP_RTU.Exceptions;

namespace ModbusTCP_RTU
{
    /// <summary>
    /// Modbus 客户端 主站
    /// </summary>
    public class ModbusClient
    {
        /// <summary>
        /// 字序
        /// </summary>
        public enum RegisterOrder
        {
            LowHigh = 0,
            HighLow = 1
        }

        private TcpClient? _tcpClient;

        /// <summary>
        /// 事务标识符（仅TCP模式）
        /// </summary>
        private uint _transactionIdentifierInternal;

        /// <summary>
        /// 事务标识符(字节形式)（仅TCP模式）
        /// </summary>
        private byte[] _transactionIdentifier = new byte[2];

        /// <summary>
        /// 协议标识符 0x0000 标识ModbusTCP协议（仅TCP模式）
        /// </summary>
        private byte[] _protocolIdentifier = new byte[2];

        /// <summary>
        /// CRC校验码（仅RTU模式）
        /// </summary>
        private byte[] _crc = new byte[2];

        /// <summary>
        /// 从当前位置开始，后边的报文长度（仅TCP模式）
        /// </summary>
        private byte[] _length = new byte[2];

        /// <summary>
        /// 功能码
        /// </summary>
        private byte _functionCode;

        /// <summary>
        /// 寄存器开始地址
        /// </summary>
        private byte[] _startingAddress = new byte[2];

        /// <summary>
        /// 寄存器数量
        /// </summary>
        private byte[] _quantity = new byte[2];

        /// <summary>
        /// UDP模式中，绑定的本地端口
        /// </summary>
        private int _portOut;

        /// <summary>
        /// 接受数据缓存区
        /// </summary>
        public byte[]? ReceiveData { get; private set; }

        /// <summary>
        /// 发送数据缓存区
        /// </summary>
        public byte[]? SendData { get; private set; }

        /// <summary>
        /// RTU模式下 串口
        /// </summary>
        private SerialPort? _serialPort;

        /// <summary>
        /// RTU模式下 奇偶校验
        /// </summary>
        private Parity _parity = Parity.Even;

        /// <summary>
        /// RTU模式下 停止位
        /// </summary>
        private StopBits _stopBits = StopBits.One;

        /// <summary>
        /// 设定或获取 RTU模式下重试次数
        /// </summary>
        public int NumberOfRetries { get; set; } = 3;

        /// <summary>
        /// RTU模式下当前重试次数
        /// </summary>
        private int _countRetries;

        public delegate void ReceiveDataChangedHandler(object sender);

        public event ReceiveDataChangedHandler? ReceiveDataChanged;

        public delegate void SendDataChangedHandler(object sender);

        public event SendDataChangedHandler? SendDataChanged;

        public delegate void ConnectedChangedHandler(object sender);

        public event ConnectedChangedHandler? ConnectedChanged;

        private NetworkStream? _stream;

        /// <summary>
        /// 创建基于TCP的客户端
        /// </summary>
        /// <param name="ipAddress">服务端IP</param>
        /// <param name="port">服务端端口</param>
        public ModbusClient(string ipAddress, int port)
        {
            IpAddress = ipAddress;
            Port = port;
        }

        /// <summary>
        /// 创建基于RTU的客户端
        /// </summary>
        /// <param name="serialPort">串口号，如："COM1"</param>
        public ModbusClient(string serialPort)
        {
            _serialPort = new SerialPort();
            _serialPort.PortName = serialPort;
            _serialPort.BaudRate = Baud;
            _serialPort.Parity = _parity;
            _serialPort.StopBits = _stopBits;
            _serialPort.WriteTimeout = 10000;
            _serialPort.ReadTimeout = ConnectionTimeout;

            _serialPort.DataReceived += DataReceivedHandler;
        }

        /// <summary>
        /// 创建一个客户端
        /// </summary>
        public ModbusClient()
        {
        }

        /// <summary>
        /// 如果使用 Modbus TCP，则与主站设备建立连接。使用Modbus RTU 时打开 COM 端口
        /// </summary>
        public void Connect()
        {
            if (_serialPort != null)
            {
                if (!_serialPort.IsOpen)
                {
                    _serialPort.BaudRate = Baud;
                    _serialPort.Parity = _parity;
                    _serialPort.StopBits = _stopBits;
                    _serialPort.WriteTimeout = 10000;
                    _serialPort.ReadTimeout = ConnectionTimeout;
                    _serialPort.Open();
                }

                ConnectedChanged?.Invoke(this);
                return;
            }

            if (!UdpFlag)
            {
                _tcpClient = new TcpClient();
                var result = _tcpClient.BeginConnect(IpAddress, Port, null, null);
                var success = result.AsyncWaitHandle.WaitOne(ConnectionTimeout);
                if (!success)
                {
                    throw new ConnectionException("连接超时");
                }

                _tcpClient.EndConnect(result);
                _stream = _tcpClient.GetStream();
                _stream.ReadTimeout = ConnectionTimeout;
            }
            else
            {
                _tcpClient = new TcpClient();
            }

            ConnectedChanged?.Invoke(this);
        }

        /// <summary>
        /// 通过 Modbus TCP，则与主站设备建立连接。
        /// </summary>
        public void Connect(string ipAddress, int port)
        {
            if (!UdpFlag)
            {
                _tcpClient = new TcpClient();
                var result = _tcpClient.BeginConnect(ipAddress, port, null, null);
                var success = result.AsyncWaitHandle.WaitOne(ConnectionTimeout);
                if (!success)
                {
                    throw new ConnectionException("连接超时");
                }

                _tcpClient.EndConnect(result);
                _stream = _tcpClient.GetStream();
                _stream.ReadTimeout = ConnectionTimeout;
            }
            else
            {
                _tcpClient = new TcpClient();
            }

            ConnectedChanged?.Invoke(this);
        }

        #region 寄存器转数据类型

        /// <summary>
        /// 将两个 ModbusRegisters 转换为浮点数 ,可设定字节顺序
        /// <code>示例：ModbusClient.ConvertRegistersToFloat(modbusClient.ReadHoldingRegisters(0,2), RegisterOrder.HighLow)</code>
        /// </summary>
        /// <param name="registers">寄存器数据</param>
        /// <param name="registerOrder">字序（低寄存器在前或高寄存器在前）</param>
        /// <returns>转化后的结果</returns>
        public static float ConvertRegistersToFloat(int[] registers,
            RegisterOrder registerOrder = RegisterOrder.LowHigh)
        {
            if (registers.Length != 2)
                throw new ArgumentException("输入数组长度无效 - 数组长度必须为 2");

            //字节交换部分
            int[] swappedRegisters = [registers[0], registers[1]];
            if (registerOrder == RegisterOrder.HighLow)
                swappedRegisters = new[] { registers[1], registers[0] };

            //开始转化
            var highRegister = swappedRegisters[1];
            var lowRegister = swappedRegisters[0];
            var highRegisterBytes = BitConverter.GetBytes(highRegister);
            var lowRegisterBytes = BitConverter.GetBytes(lowRegister);
            byte[] floatBytes =
            [
                lowRegisterBytes[0],
                lowRegisterBytes[1],
                highRegisterBytes[0],
                highRegisterBytes[1]
            ];
            return BitConverter.ToSingle(floatBytes, 0);
        }

        /// <summary>
        /// 将两个 ModbusRegisters 转换为32位有符号整数 ,可设定字节顺序
        /// <code>示例：ModbusClient.ConvertRegistersToInt(modbusClient.ReadHoldingRegisters(0,2), RegisterOrder.HighLow)</code>
        /// </summary>
        /// <param name="registers">寄存器数据</param>
        /// <param name="registerOrder">字序（低寄存器在前或高寄存器在前）</param>
        /// <returns>转化后的结果</returns>
        public static int ConvertRegistersToInt(int[] registers, RegisterOrder registerOrder = RegisterOrder.LowHigh)
        {
            if (registers.Length != 2)
                throw new ArgumentException("输入数组长度无效 - 数组长度必须为 2");

            //字节交换
            int[] swappedRegisters = [registers[0], registers[1]];
            if (registerOrder == RegisterOrder.HighLow)
                swappedRegisters = [registers[1], registers[0]];

            //开始转化
            var highRegister = swappedRegisters[1];
            var lowRegister = swappedRegisters[0];
            var highRegisterBytes = BitConverter.GetBytes(highRegister);
            var lowRegisterBytes = BitConverter.GetBytes(lowRegister);
            byte[] doubleBytes =
            [
                lowRegisterBytes[0],
                lowRegisterBytes[1],
                highRegisterBytes[0],
                highRegisterBytes[1]
            ];
            return BitConverter.ToInt32(doubleBytes, 0);
        }

        /// <summary>
        /// 将两个 ModbusRegisters 转换为32位无符号整数 ,可设定字节顺序
        /// <code>示例：ModbusClient.ConvertRegistersToUInt(modbusClient.ReadHoldingRegisters(0,2), RegisterOrder.HighLow)</code>
        /// </summary>
        /// <param name="registers">寄存器数据</param>
        /// <param name="registerOrder">字序（低寄存器在前或高寄存器在前）</param>
        /// <returns>转化后的结果</returns>
        public static uint ConvertRegistersToUInt(int[] registers, RegisterOrder registerOrder = RegisterOrder.LowHigh)
        {
            if (registers.Length != 2)
                throw new ArgumentException("输入数组长度无效 - 数组长度必须为 2");

            //字节交换
            int[] swappedRegisters = [registers[0], registers[1]];
            if (registerOrder == RegisterOrder.HighLow)
                swappedRegisters = [registers[1], registers[0]];

            //开始转化
            var highRegister = swappedRegisters[1];
            var lowRegister = swappedRegisters[0];
            var highRegisterBytes = BitConverter.GetBytes(highRegister);
            var lowRegisterBytes = BitConverter.GetBytes(lowRegister);
            byte[] doubleBytes =
            [
                lowRegisterBytes[0],
                lowRegisterBytes[1],
                highRegisterBytes[0],
                highRegisterBytes[1]
            ];
            return BitConverter.ToUInt32(doubleBytes, 0);
        }

        /// <summary>
        /// 将两个 ModbusRegisters 转换为64位有符号整数 ,可设定字节顺序
        /// <code>示例：ModbusClient.ConvertRegistersToLong(modbusClient.ReadHoldingRegisters(0,4), RegisterOrder.HighLow)</code>
        /// </summary>
        /// <param name="registers">寄存器数据</param>
        /// <param name="registerOrder">字序（低寄存器在前或高寄存器在前）</param>
        /// <returns>转化后的结果</returns>
        public static long ConvertRegistersToLong(int[] registers, RegisterOrder registerOrder = RegisterOrder.LowHigh)
        {
            if (registers.Length != 4)
                throw new ArgumentException("输入数组长度无效 - 数组长度必须为 4");

            //字节交换
            int[] swappedRegisters = [registers[0], registers[1], registers[2], registers[3]];
            if (registerOrder == RegisterOrder.HighLow)
                swappedRegisters = [registers[3], registers[2], registers[1], registers[0]];

            //开始转化
            var highRegister = swappedRegisters[3];
            var highLowRegister = swappedRegisters[2];
            var lowHighRegister = swappedRegisters[1];
            var lowRegister = swappedRegisters[0];
            var highRegisterBytes = BitConverter.GetBytes(highRegister);
            var highLowRegisterBytes = BitConverter.GetBytes(highLowRegister);
            var lowHighRegisterBytes = BitConverter.GetBytes(lowHighRegister);
            var lowRegisterBytes = BitConverter.GetBytes(lowRegister);
            byte[] longBytes =
            [
                lowRegisterBytes[0],
                lowRegisterBytes[1],
                lowHighRegisterBytes[0],
                lowHighRegisterBytes[1],
                highLowRegisterBytes[0],
                highLowRegisterBytes[1],
                highRegisterBytes[0],
                highRegisterBytes[1]
            ];
            return BitConverter.ToInt64(longBytes, 0);
        }

        /// <summary>
        /// 将两个 ModbusRegisters 转换为64位无符号整数 ,可设定字节顺序
        /// <code>示例：ModbusClient.ConvertRegistersToULong(modbusClient.ReadHoldingRegisters(0,4), RegisterOrder.HighLow)</code>
        /// </summary>
        /// <param name="registers">寄存器数据</param>
        /// <param name="registerOrder">字序（低寄存器在前或高寄存器在前）</param>
        /// <returns>转化后的结果</returns>
        public static ulong ConvertRegistersToULong(int[] registers,
            RegisterOrder registerOrder = RegisterOrder.LowHigh)
        {
            if (registers.Length != 4)
                throw new ArgumentException("输入数组长度无效 - 数组长度必须为 4");

            //字节交换
            int[] swappedRegisters = [registers[0], registers[1], registers[2], registers[3]];
            if (registerOrder == RegisterOrder.HighLow)
                swappedRegisters = [registers[3], registers[2], registers[1], registers[0]];

            //开始转化
            var highRegister = swappedRegisters[3];
            var highLowRegister = swappedRegisters[2];
            var lowHighRegister = swappedRegisters[1];
            var lowRegister = swappedRegisters[0];
            var highRegisterBytes = BitConverter.GetBytes(highRegister);
            var highLowRegisterBytes = BitConverter.GetBytes(highLowRegister);
            var lowHighRegisterBytes = BitConverter.GetBytes(lowHighRegister);
            var lowRegisterBytes = BitConverter.GetBytes(lowRegister);
            byte[] longBytes =
            [
                lowRegisterBytes[0],
                lowRegisterBytes[1],
                lowHighRegisterBytes[0],
                lowHighRegisterBytes[1],
                highLowRegisterBytes[0],
                highLowRegisterBytes[1],
                highRegisterBytes[0],
                highRegisterBytes[1]
            ];
            return BitConverter.ToUInt64(longBytes, 0);
        }

        /// <summary>
        /// 将两个 ModbusRegisters 转换为双精度浮点数 ,可设定字节顺序
        /// <code>示例：ModbusClient.ConvertRegistersToDouble(modbusClient.ReadHoldingRegisters(0,4), RegisterOrder.HighLow)</code>
        /// </summary>
        /// <param name="registers">寄存器数据</param>
        /// <param name="registerOrder">字序（低寄存器在前或高寄存器在前）</param>
        /// <returns>转化后的结果</returns>
        public static double ConvertRegistersToDouble(int[] registers,
            RegisterOrder registerOrder = RegisterOrder.LowHigh)
        {
            if (registers.Length != 4)
                throw new ArgumentException("输入数组长度无效 - 数组长度必须为 4");
            //字节交换
            int[] swappedRegisters = [registers[0], registers[1], registers[2], registers[3]];
            if (registerOrder == RegisterOrder.HighLow)
                swappedRegisters = [registers[3], registers[2], registers[1], registers[0]];
            //开始转化
            var highRegister = swappedRegisters[3];
            var highLowRegister = swappedRegisters[2];
            var lowHighRegister = swappedRegisters[1];
            var lowRegister = swappedRegisters[0];
            var highRegisterBytes = BitConverter.GetBytes(highRegister);
            var highLowRegisterBytes = BitConverter.GetBytes(highLowRegister);
            var lowHighRegisterBytes = BitConverter.GetBytes(lowHighRegister);
            var lowRegisterBytes = BitConverter.GetBytes(lowRegister);
            byte[] longBytes =
            [
                lowRegisterBytes[0],
                lowRegisterBytes[1],
                lowHighRegisterBytes[0],
                lowHighRegisterBytes[1],
                highLowRegisterBytes[0],
                highLowRegisterBytes[1],
                highRegisterBytes[0],
                highRegisterBytes[1]
            ];
            return BitConverter.ToDouble(longBytes, 0);
        }

        /// <summary>
        /// 将 ModbusRegisters 转换为字符串
        /// </summary>
        /// <param name="registers">寄存器数据</param>
        /// <param name="offset">起始位置</param>
        /// <param name="stringLength">字符串中的字符数（必须为偶数）</param>
        /// <returns>转化后的结果</returns>
        public static string ConvertRegistersToString(int[] registers, int offset, int stringLength)
        {
            var result = new byte[stringLength];
            for (var i = 0; i < stringLength / 2; i++)
            {
                var registerResult = BitConverter.GetBytes(registers[offset + i]);
                result[i * 2] = registerResult[0];
                result[i * 2 + 1] = registerResult[1];
            }

            return Encoding.Default.GetString(result);
        }

        #endregion

        #region 数据类型转寄存器

        /// <summary>
        /// 将浮点数转换为两个 ModbusRegisters
        /// <code>示例： modbusClient.WriteMultipleRegisters(0, ModbusClient.ConvertFloatToRegisters((float)1.22))；</code>
        /// </summary>
        /// <param name="floatValue">需转化的数据</param>
        /// <param name="registerOrder">字序</param>
        /// <returns>寄存器数据</returns>
        public static int[] ConvertFloatToRegisters(float floatValue,
            RegisterOrder registerOrder = RegisterOrder.LowHigh)
        {
            var floatBytes = BitConverter.GetBytes(floatValue);
            byte[] highRegisterBytes =
            [
                floatBytes[2],
                floatBytes[3],
                0,
                0
            ];
            byte[] lowRegisterBytes =
            [
                floatBytes[0],
                floatBytes[1],
                0,
                0
            ];
            int[] returnValue =
            [
                BitConverter.ToInt32(lowRegisterBytes, 0),
                BitConverter.ToInt32(highRegisterBytes, 0)
            ];

            if (registerOrder == RegisterOrder.HighLow)
                returnValue = new[] { returnValue[1], returnValue[0] };
            return returnValue;
        }

        /// <summary>
        /// 将32位有符号整数转化为两个 ModbusRegisters
        /// <code>示例： modbusClient.WriteMultipleRegisters(0, ModbusClient.ConvertIntToRegisters(32))；</code>
        /// </summary>
        /// <param name="intValue">需转化的数据</param>
        /// <param name="registerOrder">字序</param>
        /// <returns>寄存器数据</returns>
        public static int[] ConvertIntToRegisters(int intValue, RegisterOrder registerOrder = RegisterOrder.LowHigh)
        {
            var doubleBytes = BitConverter.GetBytes(intValue);
            byte[] highRegisterBytes =
            [
                doubleBytes[2],
                doubleBytes[3],
                0,
                0
            ];
            byte[] lowRegisterBytes =
            [
                doubleBytes[0],
                doubleBytes[1],
                0,
                0
            ];
            int[] returnValue =
            [
                BitConverter.ToInt32(lowRegisterBytes, 0),
                BitConverter.ToInt32(highRegisterBytes, 0)
            ];

            if (registerOrder == RegisterOrder.HighLow)
                returnValue = new[] { returnValue[1], returnValue[0] };
            return returnValue;
        }

        /// <summary>
        /// 将32位无符号整数转化为两个 ModbusRegisters
        /// <code>示例： modbusClient.WriteMultipleRegisters(0, ModbusClient.ConvertUIntToRegisters(32u))；</code>
        /// </summary>
        /// <param name="uintValue">需转化的数据</param>
        /// <param name="registerOrder">字序</param>
        /// <returns>寄存器数据</returns>
        public static int[] ConvertUIntToRegisters(uint uintValue, RegisterOrder registerOrder = RegisterOrder.LowHigh)
        {
            var doubleBytes = BitConverter.GetBytes(uintValue);
            byte[] highRegisterBytes =
            [
                doubleBytes[2],
                doubleBytes[3],
                0,
                0
            ];
            byte[] lowRegisterBytes =
            [
                doubleBytes[0],
                doubleBytes[1],
                0,
                0
            ];
            int[] returnValue =
            [
                BitConverter.ToInt32(lowRegisterBytes, 0),
                BitConverter.ToInt32(highRegisterBytes, 0)
            ];

            if (registerOrder == RegisterOrder.HighLow)
                returnValue = new[] { returnValue[1], returnValue[0] };
            return returnValue;
        }

        /// <summary>
        /// 将64位有符号整数转化为四个 ModbusRegisters
        /// <code>示例： modbusClient.WriteMultipleRegisters(0, ModbusClient.ConvertLongToRegisters(32l))；</code>
        /// </summary>
        /// <param name="longValue">需转化的数据</param>
        /// <param name="registerOrder">字序</param>
        /// <returns>寄存器数据</returns>
        public static int[] ConvertLongToRegisters(long longValue, RegisterOrder registerOrder = RegisterOrder.LowHigh)
        {
            var longBytes = BitConverter.GetBytes(longValue);
            byte[] highRegisterBytes =
            [
                longBytes[6],
                longBytes[7],
                0,
                0
            ];
            byte[] highLowRegisterBytes =
            [
                longBytes[4],
                longBytes[5],
                0,
                0
            ];
            byte[] lowHighRegisterBytes =
            [
                longBytes[2],
                longBytes[3],
                0,
                0
            ];
            byte[] lowRegisterBytes =
            [
                longBytes[0],
                longBytes[1],
                0,
                0
            ];
            int[] returnValue =
            [
                BitConverter.ToInt32(lowRegisterBytes, 0),
                BitConverter.ToInt32(lowHighRegisterBytes, 0),
                BitConverter.ToInt32(highLowRegisterBytes, 0),
                BitConverter.ToInt32(highRegisterBytes, 0)
            ];

            if (registerOrder == RegisterOrder.HighLow)
                returnValue = new[] { returnValue[3], returnValue[2], returnValue[1], returnValue[0] };
            return returnValue;
        }

        /// <summary>
        /// 将64位无符号整数转化为四个 ModbusRegisters
        /// <code>示例： modbusClient.WriteMultipleRegisters(0, ModbusClient.ConvertULongToRegisters(32l))；</code>
        /// </summary>
        /// <param name="ulongValue">需转化的数据</param>
        /// <param name="registerOrder">字序</param>
        /// <returns>寄存器数据</returns>
        public static int[] ConvertULongToRegisters(ulong ulongValue,
            RegisterOrder registerOrder = RegisterOrder.LowHigh)
        {
            var longBytes = BitConverter.GetBytes(ulongValue);
            byte[] highRegisterBytes =
            [
                longBytes[6],
                longBytes[7],
                0,
                0
            ];
            byte[] highLowRegisterBytes =
            [
                longBytes[4],
                longBytes[5],
                0,
                0
            ];
            byte[] lowHighRegisterBytes =
            [
                longBytes[2],
                longBytes[3],
                0,
                0
            ];
            byte[] lowRegisterBytes =
            [
                longBytes[0],
                longBytes[1],
                0,
                0
            ];
            int[] returnValue =
            [
                BitConverter.ToInt32(lowRegisterBytes, 0),
                BitConverter.ToInt32(lowHighRegisterBytes, 0),
                BitConverter.ToInt32(highLowRegisterBytes, 0),
                BitConverter.ToInt32(highRegisterBytes, 0)
            ];

            if (registerOrder == RegisterOrder.HighLow)
                returnValue = new[] { returnValue[3], returnValue[2], returnValue[1], returnValue[0] };
            return returnValue;
        }

        /// <summary>
        /// 将双精度浮点数转化为四个 ModbusRegisters
        /// <code>示例： modbusClient.WriteMultipleRegisters(0, ModbusClient.ConvertDoubleToRegisters(0.1))；</code>
        /// </summary>
        /// <param name="doubleValue">需转化的数据</param>
        /// <param name="registerOrder">字序</param>
        /// <returns>寄存器数据</returns>
        public static int[] ConvertDoubleToRegisters(double doubleValue,
            RegisterOrder registerOrder = RegisterOrder.LowHigh)
        {
            var doubleBytes = BitConverter.GetBytes(doubleValue);
            byte[] highRegisterBytes =
            [
                doubleBytes[6],
                doubleBytes[7],
                0,
                0
            ];
            byte[] highLowRegisterBytes =
            [
                doubleBytes[4],
                doubleBytes[5],
                0,
                0
            ];
            byte[] lowHighRegisterBytes =
            [
                doubleBytes[2],
                doubleBytes[3],
                0,
                0
            ];
            byte[] lowRegisterBytes =
            [
                doubleBytes[0],
                doubleBytes[1],
                0,
                0
            ];
            int[] returnValue =
            [
                BitConverter.ToInt32(lowRegisterBytes, 0),
                BitConverter.ToInt32(lowHighRegisterBytes, 0),
                BitConverter.ToInt32(highLowRegisterBytes, 0),
                BitConverter.ToInt32(highRegisterBytes, 0)
            ];

            if (registerOrder == RegisterOrder.HighLow)
                returnValue = new[] { returnValue[3], returnValue[2], returnValue[1], returnValue[0] };
            return returnValue;
        }


        /// <summary>
        /// 将字符串转化成 ModbusRegisters
        /// </summary>
        /// <param name="stringToConvert">将转化的字符串</param>
        /// <returns>转化后的结果</returns>
        public static int[] ConvertStringToRegisters(string stringToConvert)
        {
            var array = Encoding.ASCII.GetBytes(stringToConvert);
            var toRegisters = new int[stringToConvert.Length / 2 + stringToConvert.Length % 2];
            for (var i = 0; i < toRegisters.Length; i++)
            {
                toRegisters[i] = array[i * 2];
                if (i * 2 + 1 < array.Length)
                {
                    toRegisters[i] |= array[i * 2 + 1] << 8;
                }
            }

            return toRegisters;
        }

        #endregion

        /// <summary>
        /// 计算 Modbus-RTU 的 CRC16
        /// </summary>
        /// <param name="data">要发送的字节缓冲区</param>
        /// <param name="numberOfBytes">计算 CRC 的字节数</param>
        /// <param name="startByte">缓冲区中开始计算 CRC 的第一个字节</param>
        public static ushort CalculateCrc(byte[] data, ushort numberOfBytes, int startByte)
        {
            byte[] auchCrcHi =
            [
                0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81,
                0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0,
                0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01,
                0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41,
                0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81,
                0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0,
                0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01,
                0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40,
                0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81,
                0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0,
                0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01,
                0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41,
                0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81,
                0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0,
                0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01,
                0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81, 0x40, 0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41,
                0x00, 0xC1, 0x81, 0x40, 0x01, 0xC0, 0x80, 0x41, 0x01, 0xC0, 0x80, 0x41, 0x00, 0xC1, 0x81,
                0x40
            ];

            byte[] auchCrcLo =
            [
                0x00, 0xC0, 0xC1, 0x01, 0xC3, 0x03, 0x02, 0xC2, 0xC6, 0x06, 0x07, 0xC7, 0x05, 0xC5, 0xC4,
                0x04, 0xCC, 0x0C, 0x0D, 0xCD, 0x0F, 0xCF, 0xCE, 0x0E, 0x0A, 0xCA, 0xCB, 0x0B, 0xC9, 0x09,
                0x08, 0xC8, 0xD8, 0x18, 0x19, 0xD9, 0x1B, 0xDB, 0xDA, 0x1A, 0x1E, 0xDE, 0xDF, 0x1F, 0xDD,
                0x1D, 0x1C, 0xDC, 0x14, 0xD4, 0xD5, 0x15, 0xD7, 0x17, 0x16, 0xD6, 0xD2, 0x12, 0x13, 0xD3,
                0x11, 0xD1, 0xD0, 0x10, 0xF0, 0x30, 0x31, 0xF1, 0x33, 0xF3, 0xF2, 0x32, 0x36, 0xF6, 0xF7,
                0x37, 0xF5, 0x35, 0x34, 0xF4, 0x3C, 0xFC, 0xFD, 0x3D, 0xFF, 0x3F, 0x3E, 0xFE, 0xFA, 0x3A,
                0x3B, 0xFB, 0x39, 0xF9, 0xF8, 0x38, 0x28, 0xE8, 0xE9, 0x29, 0xEB, 0x2B, 0x2A, 0xEA, 0xEE,
                0x2E, 0x2F, 0xEF, 0x2D, 0xED, 0xEC, 0x2C, 0xE4, 0x24, 0x25, 0xE5, 0x27, 0xE7, 0xE6, 0x26,
                0x22, 0xE2, 0xE3, 0x23, 0xE1, 0x21, 0x20, 0xE0, 0xA0, 0x60, 0x61, 0xA1, 0x63, 0xA3, 0xA2,
                0x62, 0x66, 0xA6, 0xA7, 0x67, 0xA5, 0x65, 0x64, 0xA4, 0x6C, 0xAC, 0xAD, 0x6D, 0xAF, 0x6F,
                0x6E, 0xAE, 0xAA, 0x6A, 0x6B, 0xAB, 0x69, 0xA9, 0xA8, 0x68, 0x78, 0xB8, 0xB9, 0x79, 0xBB,
                0x7B, 0x7A, 0xBA, 0xBE, 0x7E, 0x7F, 0xBF, 0x7D, 0xBD, 0xBC, 0x7C, 0xB4, 0x74, 0x75, 0xB5,
                0x77, 0xB7, 0xB6, 0x76, 0x72, 0xB2, 0xB3, 0x73, 0xB1, 0x71, 0x70, 0xB0, 0x50, 0x90, 0x91,
                0x51, 0x93, 0x53, 0x52, 0x92, 0x96, 0x56, 0x57, 0x97, 0x55, 0x95, 0x94, 0x54, 0x9C, 0x5C,
                0x5D, 0x9D, 0x5F, 0x9F, 0x9E, 0x5E, 0x5A, 0x9A, 0x9B, 0x5B, 0x99, 0x59, 0x58, 0x98, 0x88,
                0x48, 0x49, 0x89, 0x4B, 0x8B, 0x8A, 0x4A, 0x4E, 0x8E, 0x8F, 0x4F, 0x8D, 0x4D, 0x4C, 0x8C,
                0x44, 0x84, 0x85, 0x45, 0x87, 0x47, 0x46, 0x86, 0x82, 0x42, 0x43, 0x83, 0x41, 0x81, 0x80,
                0x40
            ];
            var usDataLen = numberOfBytes;
            byte uchCrcHi = 0xFF;
            byte uchCrcLo = 0xFF;
            var i = 0;
            while (usDataLen > 0)
            {
                usDataLen--;
                if ((i + startByte) < data.Length)
                {
                    var uIndex = uchCrcLo ^ data[i + startByte];
                    uchCrcLo = (byte)(uchCrcHi ^ auchCrcHi[uIndex]);
                    uchCrcHi = auchCrcLo[uIndex];
                }

                i++;
            }

            return (ushort)(uchCrcHi << 8 | uchCrcLo);
        }

        /// <summary>
        /// RTU模式下，数据接收完成标志
        /// </summary>
        private bool _dataReceived;

        /// <summary>
        /// RTU模式下，正在接收数据状态
        /// </summary>
        private bool _receiveActive;

        /// <summary>
        /// RTU模式下，接受数据缓存区
        /// </summary>
        private byte[] _readBuffer = new byte[256];

        /// <summary>
        /// 数据长度。用于发送成功后，指定要读取多少数据
        /// </summary>
        private int _bytesToRead;

        /// <summary>
        /// RTU模式下接收到数据时触发
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void DataReceivedHandler(object sender,
            SerialDataReceivedEventArgs e)
        {
            if (_serialPort == null)
                return;

            _serialPort.DataReceived -= DataReceivedHandler;
            _receiveActive = true;
            const long ticksWait = TimeSpan.TicksPerMillisecond * 2000;

            var sp = (SerialPort)sender;
            if (_bytesToRead == 0)
            {
                sp.DiscardInBuffer();
                _receiveActive = false;
                _serialPort.DataReceived += DataReceivedHandler;
                return;
            }

            _readBuffer = new byte[256];
            var actualPositionToRead = 0;
            var dateTimeLastRead = DateTime.Now;
            do
            {
                try
                {
                    dateTimeLastRead = DateTime.Now;
                    while ((sp.BytesToRead) == 0)
                    {
                        Thread.Sleep(10);
                        if ((DateTime.Now.Ticks - dateTimeLastRead.Ticks) > ticksWait)
                            break;
                    }

                    var numByte = sp.BytesToRead;
                    var rxByteArray = new byte[numByte];
                    sp.Read(rxByteArray, 0, numByte);
                    Array.Copy(rxByteArray, 0, _readBuffer, actualPositionToRead,
                        (actualPositionToRead + rxByteArray.Length) <= _bytesToRead
                            ? rxByteArray.Length
                            : _bytesToRead - actualPositionToRead);

                    actualPositionToRead += rxByteArray.Length;
                }
                catch (Exception)
                {
                    // ignored
                }

                if (_bytesToRead <= actualPositionToRead)
                    break;

                if (DetectValidModbusFrame(_readBuffer,
                        (actualPositionToRead < _readBuffer.Length) ? actualPositionToRead : _readBuffer.Length) |
                    _bytesToRead <= actualPositionToRead)
                    break;
            } while ((DateTime.Now.Ticks - dateTimeLastRead.Ticks) < ticksWait);
            //数据接收完成

            ReceiveData = new byte[actualPositionToRead];
            Array.Copy(_readBuffer, 0, ReceiveData, 0,
                (actualPositionToRead < _readBuffer.Length) ? actualPositionToRead : _readBuffer.Length);

            _bytesToRead = 0;
            _dataReceived = true;
            _receiveActive = false;
            _serialPort.DataReceived += DataReceivedHandler;
            ReceiveDataChanged?.Invoke(this);
        }

        /// <summary>
        /// 检测 Modbus 帧 是否有效
        /// </summary>
        /// <param name="readBuffer"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        public static bool DetectValidModbusFrame(byte[] readBuffer, int length)
        {
            // 最小长度 6 字节
            if (length < 6)
                return false;
            //从站地址
            if ((readBuffer[0] < 1) | (readBuffer[0] > 247))
                return false;
            //CRC 校验
            var crc = BitConverter.GetBytes(CalculateCrc(readBuffer, (ushort)(length - 2), 0));
            return !(crc[0] != readBuffer[length - 2] | crc[1] != readBuffer[length - 1]);
        }

        /// <summary>
        /// 检查故障代码
        /// </summary>
        /// <param name="code">故障代码</param>
        /// <exception cref="ModbusException">故障抛出</exception>
        private static void CheckErrorCode(byte code)
        {
            if (code == 0x00)
                return;

            throw code switch
            {
                0x01 => new FunctionCodeNotSupportedException("非法功能"),
                0x02 => new StartingAddressInvalidException("非法数据地址"),
                0x03 => new QuantityInvalidException("非法数据值"),
                0x04 => new ModbusException("从站设备故障"),
                0x05 => new ModbusException("确认"),
                0x06 => new ModbusException("从属设备忙"),
                0x0A => new ModbusException("不可用网关路径"),
                0x0B => new ModbusException("网关目标设备响应失败"),
                _ => new ModbusException("未知错误")
            };
        }

        /// <summary>
        /// 读离散输入状态 (FC2) 。
        /// </summary>
        /// <param name="startingAddress">起始地址</param>
        /// <param name="quantity">读取数量</param>
        /// <returns>读取结果</returns>
        public bool[] ReadDiscreteInputs(int startingAddress, int quantity)
        {
            _transactionIdentifierInternal++;
            if (_serialPort is { IsOpen: false })
            {
                throw new SerialPortNotOpenedException("串行端口未打开");
            }

            if (_tcpClient == null & !UdpFlag & _serialPort == null)
            {
                throw new ConnectionException("连接错误");
            }

            if (startingAddress > 65535 | quantity > 2000)
            {
                throw new ArgumentException("起始地址必须为 0 - 65535；数量必须为 0 - 2000");
            }

            _transactionIdentifier = BitConverter.GetBytes(_transactionIdentifierInternal);
            _protocolIdentifier = BitConverter.GetBytes(0x0000);
            _length = BitConverter.GetBytes(0x0006);
            _functionCode = 0x02;
            _startingAddress = BitConverter.GetBytes(startingAddress);
            _quantity = BitConverter.GetBytes(quantity);
            byte[] data =
            [
                _transactionIdentifier[1],
                _transactionIdentifier[0],
                _protocolIdentifier[1],
                _protocolIdentifier[0],
                _length[1],
                _length[0],
                UnitIdentifier,
                _functionCode,
                _startingAddress[1],
                _startingAddress[0],
                _quantity[1],
                _quantity[0],
                _crc[0],
                _crc[1]
            ];
            _crc = BitConverter.GetBytes(CalculateCrc(data, 6, 6));
            data[12] = _crc[0];
            data[13] = _crc[1];

            if (_serialPort != null)
            {
                _dataReceived = false;
                if (quantity % 8 == 0)
                    _bytesToRead = 5 + quantity / 8;
                else
                    _bytesToRead = 6 + quantity / 8;
                //RTU模式没有事务标识符，协议标识符和长度，共 6 字节
                _serialPort.Write(data, 6, 8);

                if (SendDataChanged != null)
                {
                    SendData = new byte[8];
                    Array.Copy(data, 6, SendData, 0, 8);
                    SendDataChanged(this);
                }

                data = new byte[2100];
                _readBuffer = new byte[256];
                var dateTimeSend = DateTime.Now;
                byte receivedUnitIdentifier = 0xFF; //接收到的从站地址

                // 10000L = 1 ms，接收数据
                while (receivedUnitIdentifier != UnitIdentifier & !(DateTime.Now.Ticks - dateTimeSend.Ticks >
                                                                    TimeSpan.TicksPerMillisecond *
                                                                    ConnectionTimeout))
                {
                    //等待数据接收完成
                    while (_dataReceived == false & !(DateTime.Now.Ticks - dateTimeSend.Ticks >
                                                      TimeSpan.TicksPerMillisecond * ConnectionTimeout))
                        Thread.Sleep(1);

                    data = new byte[2100];
                    Array.Copy(_readBuffer, 0, data, 6, _readBuffer.Length);
                    receivedUnitIdentifier = data[6];
                }

                //检查从站地址是否匹配
                if (receivedUnitIdentifier != UnitIdentifier)
                    data = new byte[2100];
                else
                    _countRetries = 0;
            }
            else if (_tcpClient != null && _tcpClient.Client.Connected | UdpFlag)
            {
                if (UdpFlag)
                {
                    var udpClient = new UdpClient();
                    var endPoint = new IPEndPoint(IPAddress.Parse(IpAddress), Port);
                    udpClient.Send(data, data.Length - 2, endPoint);
                    if (udpClient.Client.LocalEndPoint is IPEndPoint point)
                        _portOut = point.Port;
                    else
                    {
                        udpClient.Close();
                        throw new ConnectionException("连接错误，未找到本地端口");
                    }

                    udpClient.Client.ReceiveTimeout = 5000;
                    endPoint = new IPEndPoint(IPAddress.Parse(IpAddress), _portOut);
                    data = udpClient.Receive(ref endPoint);
                }
                else
                {
                    if (_stream != null)
                    {
                        _stream.Write(data, 0, data.Length - 2);

                        if (SendDataChanged != null)
                        {
                            SendData = new byte[data.Length - 2];
                            Array.Copy(data, 0, SendData, 0, data.Length - 2);
                            SendDataChanged(this);
                        }

                        data = new byte[2100];
                        var numberOfBytes = _stream.Read(data, 0, data.Length);
                        if (ReceiveDataChanged != null)
                        {
                            ReceiveData = new byte[numberOfBytes];
                            Array.Copy(data, 0, ReceiveData, 0, numberOfBytes);
                            ReceiveDataChanged(this);
                        }
                    }
                    else
                    {
                        throw new ConnectionException("连接错误，stream 为空");
                    }
                }
            }
            else
            {
                throw new ConnectionException("连接错误，未找到连接");
            }

            CheckErrorCode(data[8]);

            if (_serialPort != null)
            {
                _crc = BitConverter.GetBytes(CalculateCrc(data, (ushort)(data[8] + 3), 6));
                if ((_crc[0] != data[data[8] + 9] | _crc[1] != data[data[8] + 10]) & _dataReceived)
                {
                    if (NumberOfRetries <= _countRetries)
                    {
                        _countRetries = 0;
                        throw new CRCCheckFailedException("响应 CRC 校验失败");
                    }

                    //失败重试
                    _countRetries++;
                    return ReadDiscreteInputs(startingAddress, quantity);
                }

                if (!_dataReceived)
                {
                    if (NumberOfRetries <= _countRetries)
                    {
                        _countRetries = 0;
                        throw new TimeoutException("Modbus 从站无响应");
                    }

                    //失败重试
                    _countRetries++;
                    return ReadDiscreteInputs(startingAddress, quantity);
                }
            }

            var response = new bool[quantity];
            for (var i = 0; i < quantity; i++)
            {
                int intData = data[9 + i / 8];
                var mask = Convert.ToInt32(Math.Pow(2, i % 8));
                response[i] = Convert.ToBoolean((intData & mask) / mask);
            }

            return response;
        }


        /// <summary>
        /// 读线圈状态 (FC1) 。
        /// </summary>
        /// <param name="startingAddress">起始地址</param>
        /// <param name="quantity">读取数量</param>
        /// <returns>读取结果</returns>
        public bool[] ReadCoils(int startingAddress, int quantity)
        {
            _transactionIdentifierInternal++;
            if (_serialPort is { IsOpen: false })
            {
                throw new SerialPortNotOpenedException("串行端口未打开");
            }

            if (_tcpClient == null & !UdpFlag & _serialPort == null)
            {
                throw new ConnectionException("连接错误");
            }

            if (startingAddress > 65535 | quantity > 2000)
            {
                throw new ArgumentException("起始地址必须为 0 - 65535；数量必须为 0 - 2000");
            }

            _transactionIdentifier = BitConverter.GetBytes(_transactionIdentifierInternal);
            _protocolIdentifier = BitConverter.GetBytes(0x0000);
            _length = BitConverter.GetBytes(0x0006);
            _functionCode = 0x01;
            _startingAddress = BitConverter.GetBytes(startingAddress);
            _quantity = BitConverter.GetBytes(quantity);
            byte[] data =
            [
                _transactionIdentifier[1],
                _transactionIdentifier[0],
                _protocolIdentifier[1],
                _protocolIdentifier[0],
                _length[1],
                _length[0],
                UnitIdentifier,
                _functionCode,
                _startingAddress[1],
                _startingAddress[0],
                _quantity[1],
                _quantity[0],
                _crc[0],
                _crc[1]
            ];

            _crc = BitConverter.GetBytes(CalculateCrc(data, 6, 6));
            data[12] = _crc[0];
            data[13] = _crc[1];
            if (_serialPort != null)
            {
                _dataReceived = false;
                if (quantity % 8 == 0)
                    _bytesToRead = 5 + quantity / 8;
                else
                    _bytesToRead = 6 + quantity / 8;
                //RTU模式没有事务标识符，协议标识符和长度，共 6 字节
                _serialPort.Write(data, 6, 8);

                if (SendDataChanged != null)
                {
                    SendData = new byte[8];
                    Array.Copy(data, 6, SendData, 0, 8);
                    SendDataChanged(this);
                }

                data = new byte[2100];
                _readBuffer = new byte[256];
                var dateTimeSend = DateTime.Now;
                byte receivedUnitIdentifier = 0xFF; //接收到的从站地址

                // 10000L = 1 ms，接收数据
                while (receivedUnitIdentifier != UnitIdentifier & !((DateTime.Now.Ticks - dateTimeSend.Ticks) >
                                                                    TimeSpan.TicksPerMillisecond *
                                                                    ConnectionTimeout))
                {
                    //等待数据接收完成
                    while (_dataReceived == false & !((DateTime.Now.Ticks - dateTimeSend.Ticks) >
                                                      TimeSpan.TicksPerMillisecond * ConnectionTimeout))
                        Thread.Sleep(1);
                    data = new byte[2100];

                    Array.Copy(_readBuffer, 0, data, 6, _readBuffer.Length);
                    receivedUnitIdentifier = data[6];
                }

                //检查从站地址是否匹配
                if (receivedUnitIdentifier != UnitIdentifier)
                    data = new byte[2100];
                else
                    _countRetries = 0;
            }
            else if (_tcpClient != null && _tcpClient.Client.Connected | UdpFlag)
            {
                if (UdpFlag)
                {
                    var udpClient = new UdpClient();
                    var endPoint = new IPEndPoint(IPAddress.Parse(IpAddress), Port);
                    udpClient.Send(data, data.Length - 2, endPoint);

                    if (udpClient.Client.LocalEndPoint is IPEndPoint point)
                        _portOut = point.Port;
                    else
                    {
                        udpClient.Close();
                        throw new ConnectionException("连接错误，未找到本地端口");
                    }

                    udpClient.Client.ReceiveTimeout = 5000;
                    endPoint = new IPEndPoint(IPAddress.Parse(IpAddress), _portOut);
                    data = udpClient.Receive(ref endPoint);
                }
                else
                {
                    if (_stream != null)
                    {
                        _stream.Write(data, 0, data.Length - 2);

                        if (SendDataChanged != null)
                        {
                            SendData = new byte[data.Length - 2];
                            Array.Copy(data, 0, SendData, 0, data.Length - 2);
                            SendDataChanged(this);
                        }

                        data = new byte[2100];
                        var numberOfBytes = _stream.Read(data, 0, data.Length);
                        if (ReceiveDataChanged != null)
                        {
                            ReceiveData = new byte[numberOfBytes];
                            Array.Copy(data, 0, ReceiveData, 0, numberOfBytes);
                            ReceiveDataChanged(this);
                        }
                    }
                    else
                    {
                        throw new ConnectionException("连接错误，stream 为空");
                    }
                }
            }
            else
            {
                throw new ConnectionException("连接错误，未找到连接");
            }

            CheckErrorCode(data[8]);

            if (_serialPort != null)
            {
                _crc = BitConverter.GetBytes(CalculateCrc(data, (ushort)(data[8] + 3), 6));
                if ((_crc[0] != data[data[8] + 9] | _crc[1] != data[data[8] + 10]) & _dataReceived)
                {
                    if (NumberOfRetries <= _countRetries)
                    {
                        _countRetries = 0;
                        throw new CRCCheckFailedException("响应 CRC 校验失败");
                    }

                    _countRetries++;
                    return ReadCoils(startingAddress, quantity);
                }

                if (!_dataReceived)
                {
                    if (NumberOfRetries <= _countRetries)
                    {
                        _countRetries = 0;
                        throw new TimeoutException("Modbus 从站无响应");
                    }

                    _countRetries++;
                    return ReadCoils(startingAddress, quantity);
                }
            }

            var response = new bool[quantity];
            for (var i = 0; i < quantity; i++)
            {
                int intData = data[9 + i / 8];
                var mask = Convert.ToInt32(Math.Pow(2, i % 8));
                response[i] = Convert.ToBoolean((intData & mask) / mask);
            }

            return response;
        }


        /// <summary>
        /// 读保持寄存器内容 (FC3).
        /// </summary>
        /// <param name="startingAddress">起始地址</param>
        /// <param name="quantity">读取数量</param>
        /// <returns>读取结果</returns>
        public int[] ReadHoldingRegisters(int startingAddress, int quantity)
        {
            _transactionIdentifierInternal++;
            if (_serialPort is { IsOpen: false })
            {
                throw new SerialPortNotOpenedException("串行端口未打开");
            }

            if (_tcpClient == null & !UdpFlag & _serialPort == null)
            {
                throw new ConnectionException("连接错误");
            }

            if (startingAddress > 65535 | quantity > 125)
            {
                throw new ArgumentException("起始地址必须为 0 - 65535；数量必须为 0 - 125");
            }

            _transactionIdentifier = BitConverter.GetBytes(_transactionIdentifierInternal);
            _protocolIdentifier = BitConverter.GetBytes(0x0000);
            _length = BitConverter.GetBytes(0x0006);
            _functionCode = 0x03;
            _startingAddress = BitConverter.GetBytes(startingAddress);
            _quantity = BitConverter.GetBytes(quantity);
            byte[] data =
            [
                _transactionIdentifier[1],
                _transactionIdentifier[0],
                _protocolIdentifier[1],
                _protocolIdentifier[0],
                _length[1],
                _length[0],
                UnitIdentifier,
                _functionCode,
                _startingAddress[1],
                _startingAddress[0],
                _quantity[1],
                _quantity[0],
                _crc[0],
                _crc[1]
            ];
            _crc = BitConverter.GetBytes(CalculateCrc(data, 6, 6));
            data[12] = _crc[0];
            data[13] = _crc[1];
            if (_serialPort != null)
            {
                _dataReceived = false;
                _bytesToRead = 5 + 2 * quantity;
                //RTU模式没有事务标识符，协议标识符和长度，共 6 字节
                _serialPort.Write(data, 6, 8);
                if (SendDataChanged != null)
                {
                    SendData = new byte[8];
                    Array.Copy(data, 6, SendData, 0, 8);
                    SendDataChanged(this);
                }

                data = new byte[2100];
                _readBuffer = new byte[256];

                var dateTimeSend = DateTime.Now;
                byte receivedUnitIdentifier = 0xFF; //接收到的从站地址

                // 10000L = 1 ms，接收数据
                while (receivedUnitIdentifier != UnitIdentifier & !((DateTime.Now.Ticks - dateTimeSend.Ticks) >
                                                                    TimeSpan.TicksPerMillisecond *
                                                                    ConnectionTimeout))
                {
                    //等待数据接收完成
                    while (_dataReceived == false & !((DateTime.Now.Ticks - dateTimeSend.Ticks) >
                                                      TimeSpan.TicksPerMillisecond * ConnectionTimeout))
                        Thread.Sleep(1);
                    data = new byte[2100];
                    Array.Copy(_readBuffer, 0, data, 6, _readBuffer.Length);

                    receivedUnitIdentifier = data[6];
                }

                //检查从站地址是否匹配
                if (receivedUnitIdentifier != UnitIdentifier)
                    data = new byte[2100];
                else
                    _countRetries = 0;
            }
            else if (_tcpClient != null && _tcpClient.Client.Connected | UdpFlag)
            {
                if (UdpFlag)
                {
                    var udpClient = new UdpClient();
                    var endPoint = new IPEndPoint(IPAddress.Parse(IpAddress), Port);
                    udpClient.Send(data, data.Length - 2, endPoint);

                    if (udpClient.Client.LocalEndPoint is IPEndPoint point)
                        _portOut = point.Port;
                    else
                    {
                        udpClient.Close();
                        throw new ConnectionException("连接错误，未找到本地端口");
                    }

                    udpClient.Client.ReceiveTimeout = 5000;
                    endPoint = new IPEndPoint(IPAddress.Parse(IpAddress), _portOut);
                    data = udpClient.Receive(ref endPoint);
                }
                else
                {
                    if (_stream != null)
                    {
                        _stream.Write(data, 0, data.Length - 2);
                        if (SendDataChanged != null)
                        {
                            SendData = new byte[data.Length - 2];
                            Array.Copy(data, 0, SendData, 0, data.Length - 2);
                            SendDataChanged(this);
                        }

                        data = new byte[2100];
                        var numberOfBytes = _stream.Read(data, 0, data.Length);
                        if (ReceiveDataChanged != null)
                        {
                            ReceiveData = new byte[numberOfBytes];
                            Array.Copy(data, 0, ReceiveData, 0, numberOfBytes);
                            ReceiveDataChanged(this);
                        }
                    }
                    else
                    {
                        throw new ConnectionException("连接错误，stream 为空");
                    }
                }
            }
            else
            {
                throw new ConnectionException("连接错误，未找到连接");
            }

            CheckErrorCode(data[8]);

            if (_serialPort != null)
            {
                _crc = BitConverter.GetBytes(CalculateCrc(data, (ushort)(data[8] + 3), 6));
                if ((_crc[0] != data[data[8] + 9] | _crc[1] != data[data[8] + 10]) & _dataReceived)
                {
                    if (NumberOfRetries <= _countRetries)
                    {
                        _countRetries = 0;
                        throw new CRCCheckFailedException("响应 CRC 校验失败");
                    }

                    //失败重试
                    _countRetries++;
                    return ReadHoldingRegisters(startingAddress, quantity);
                }

                if (!_dataReceived)
                {
                    if (NumberOfRetries <= _countRetries)
                    {
                        _countRetries = 0;
                        throw new TimeoutException("Modbus 从站无响应");
                    }

                    //失败重试
                    _countRetries++;
                    return ReadHoldingRegisters(startingAddress, quantity);
                }
            }

            var response = new int[quantity];
            for (var i = 0; i < quantity; i++)
            {
                var highByte = data[9 + i * 2];
                var lowByte = data[9 + i * 2 + 1];

                data[9 + i * 2] = lowByte;
                data[9 + i * 2 + 1] = highByte;

                response[i] = BitConverter.ToInt16(data, (9 + i * 2));
            }

            return response;
        }


        /// <summary>
        /// 读输入寄存器内容 (FC4).
        /// </summary>
        /// <param name="startingAddress">起始地址</param>
        /// <param name="quantity">读取数量</param>
        /// <returns>读取结果</returns>
        public int[] ReadInputRegisters(int startingAddress, int quantity)
        {
            _transactionIdentifierInternal++;
            if (_serialPort is { IsOpen: false })
            {
                throw new SerialPortNotOpenedException("串行端口未打开");
            }

            if (_tcpClient == null & !UdpFlag & _serialPort == null)
            {
                throw new ConnectionException("连接错误");
            }

            if (startingAddress > 65535 | quantity > 125)
            {
                throw new ArgumentException("起始地址必须为 0 - 65535；数量必须为 0 - 125");
            }

            _transactionIdentifier = BitConverter.GetBytes(_transactionIdentifierInternal);
            _protocolIdentifier = BitConverter.GetBytes(0x0000);
            _length = BitConverter.GetBytes(0x0006);
            _functionCode = 0x04;
            _startingAddress = BitConverter.GetBytes(startingAddress);
            _quantity = BitConverter.GetBytes(quantity);
            byte[] data =
            [
                _transactionIdentifier[1],
                _transactionIdentifier[0],
                _protocolIdentifier[1],
                _protocolIdentifier[0],
                _length[1],
                _length[0],
                UnitIdentifier,
                _functionCode,
                _startingAddress[1],
                _startingAddress[0],
                _quantity[1],
                _quantity[0],
                _crc[0],
                _crc[1]
            ];
            _crc = BitConverter.GetBytes(CalculateCrc(data, 6, 6));
            data[12] = _crc[0];
            data[13] = _crc[1];
            if (_serialPort != null)
            {
                _dataReceived = false;
                _bytesToRead = 5 + 2 * quantity;


                //RTU模式没有事务标识符，协议标识符和长度，共 6 字节
                _serialPort.Write(data, 6, 8);
                if (SendDataChanged != null)
                {
                    SendData = new byte[8];
                    Array.Copy(data, 6, SendData, 0, 8);
                    SendDataChanged(this);
                }

                data = new byte[2100];
                _readBuffer = new byte[256];
                var dateTimeSend = DateTime.Now;
                byte receivedUnitIdentifier = 0xFF; //接收到的从站地址

                // 10000L = 1 ms，接收数据
                while (receivedUnitIdentifier != UnitIdentifier & !(DateTime.Now.Ticks - dateTimeSend.Ticks >
                                                                    TimeSpan.TicksPerMillisecond *
                                                                    ConnectionTimeout))
                {
                    //等待数据接收完成
                    while (_dataReceived == false & !(DateTime.Now.Ticks - dateTimeSend.Ticks >
                                                      TimeSpan.TicksPerMillisecond * ConnectionTimeout))
                        Thread.Sleep(1);

                    data = new byte[2100];
                    Array.Copy(_readBuffer, 0, data, 6, _readBuffer.Length);
                    receivedUnitIdentifier = data[6];
                }

                //检查从站地址是否匹配
                if (receivedUnitIdentifier != UnitIdentifier)
                    data = new byte[2100];
                else
                    _countRetries = 0;
            }
            else if (_tcpClient != null && _tcpClient.Client.Connected | UdpFlag)
            {
                if (UdpFlag)
                {
                    var udpClient = new UdpClient();
                    var endPoint = new IPEndPoint(IPAddress.Parse(IpAddress), Port);
                    udpClient.Send(data, data.Length - 2, endPoint);

                    if (udpClient.Client.LocalEndPoint is IPEndPoint point)
                        _portOut = point.Port;
                    else
                    {
                        udpClient.Close();
                        throw new ConnectionException("连接错误，未找到本地端口");
                    }

                    udpClient.Client.ReceiveTimeout = 5000;
                    endPoint = new IPEndPoint(IPAddress.Parse(IpAddress), _portOut);
                    data = udpClient.Receive(ref endPoint);
                }
                else
                {
                    if (_stream != null)
                    {
                        _stream.Write(data, 0, data.Length - 2);

                        if (SendDataChanged != null)
                        {
                            SendData = new byte[data.Length - 2];
                            Array.Copy(data, 0, SendData, 0, data.Length - 2);
                            SendDataChanged(this);
                        }

                        data = new byte[2100];
                        var numberOfBytes = _stream.Read(data, 0, data.Length);
                        if (ReceiveDataChanged != null)
                        {
                            ReceiveData = new byte[numberOfBytes];
                            Array.Copy(data, 0, ReceiveData, 0, numberOfBytes);
                            ReceiveDataChanged(this);
                        }
                    }
                    else
                    {
                        throw new ConnectionException("连接错误，stream 为空");
                    }
                }
            }
            else
            {
                throw new ConnectionException("连接错误，未找到连接");
            }

            CheckErrorCode(data[8]);

            if (_serialPort != null)
            {
                _crc = BitConverter.GetBytes(CalculateCrc(data, (ushort)(data[8] + 3), 6));
                if ((_crc[0] != data[data[8] + 9] | _crc[1] != data[data[8] + 10]) & _dataReceived)
                {
                    if (NumberOfRetries <= _countRetries)
                    {
                        _countRetries = 0;
                        throw new CRCCheckFailedException("响应 CRC 校验失败");
                    }

                    _countRetries++;
                    return ReadInputRegisters(startingAddress, quantity);
                }

                if (!_dataReceived)
                {
                    if (NumberOfRetries <= _countRetries)
                    {
                        _countRetries = 0;
                        throw new TimeoutException("Modbus 从站无响应");
                    }

                    _countRetries++;
                    return ReadInputRegisters(startingAddress, quantity);
                }
            }

            var response = new int[quantity];
            for (var i = 0; i < quantity; i++)
            {
                var highByte = data[9 + i * 2];
                var lowByte = data[9 + i * 2 + 1];

                data[9 + i * 2] = lowByte;
                data[9 + i * 2 + 1] = highByte;

                response[i] = BitConverter.ToInt16(data, (9 + i * 2));
            }

            return response;
        }


        /// <summary>
        /// 写单个线圈 (FC5).
        /// </summary>
        /// <param name="startingAddress">起始地址</param>
        /// <param name="value">值</param>
        public void WriteSingleCoil(int startingAddress, bool value)
        {
            _transactionIdentifierInternal++;
            if (_serialPort is { IsOpen: false })
            {
                throw new SerialPortNotOpenedException("串行端口未打开");
            }

            if (_tcpClient == null & !UdpFlag & _serialPort == null)
            {
                throw new ConnectionException("连接错误");
            }

            _transactionIdentifier = BitConverter.GetBytes(_transactionIdentifierInternal);
            _protocolIdentifier = BitConverter.GetBytes(0x0000);
            _length = BitConverter.GetBytes(0x0006);
            _functionCode = 0x05;
            _startingAddress = BitConverter.GetBytes(startingAddress);
            var coilValue = BitConverter.GetBytes(value ? 0xFF00 : 0x0000);

            byte[] data =
            [
                _transactionIdentifier[1],
                _transactionIdentifier[0],
                _protocolIdentifier[1],
                _protocolIdentifier[0],
                _length[1],
                _length[0],
                UnitIdentifier,
                _functionCode,
                _startingAddress[1],
                _startingAddress[0],
                coilValue[1],
                coilValue[0],
                _crc[0],
                _crc[1]
            ];
            _crc = BitConverter.GetBytes(CalculateCrc(data, 6, 6));
            data[12] = _crc[0];
            data[13] = _crc[1];
            if (_serialPort != null)
            {
                _dataReceived = false;
                _bytesToRead = 8;
                _serialPort.Write(data, 6, 8);

                if (SendDataChanged != null)
                {
                    SendData = new byte[8];
                    Array.Copy(data, 6, SendData, 0, 8);
                    SendDataChanged(this);
                }

                data = new byte[2100];
                _readBuffer = new byte[256];
                var dateTimeSend = DateTime.Now;
                byte receivedUnitIdentifier = 0xFF; //接收到的从站地址

                // 10000L = 1 ms，接收数据
                while (receivedUnitIdentifier != UnitIdentifier & !(DateTime.Now.Ticks - dateTimeSend.Ticks >
                                                                    TimeSpan.TicksPerMillisecond *
                                                                    ConnectionTimeout))
                {
                    //等待数据接收完成
                    while (_dataReceived == false & !(DateTime.Now.Ticks - dateTimeSend.Ticks >
                                                      TimeSpan.TicksPerMillisecond * ConnectionTimeout))
                        Thread.Sleep(1);

                    data = new byte[2100];
                    Array.Copy(_readBuffer, 0, data, 6, _readBuffer.Length);
                    receivedUnitIdentifier = data[6];
                }

                //检查从站地址是否匹配
                if (receivedUnitIdentifier != UnitIdentifier)
                    data = new byte[2100];
                else
                    _countRetries = 0;
            }
            else if (_tcpClient != null && _tcpClient.Client.Connected | UdpFlag)
            {
                if (UdpFlag)
                {
                    var udpClient = new UdpClient();
                    var endPoint = new IPEndPoint(IPAddress.Parse(IpAddress), Port);
                    udpClient.Send(data, data.Length - 2, endPoint);

                    if (udpClient.Client.LocalEndPoint is IPEndPoint point)
                        _portOut = point.Port;
                    else
                    {
                        udpClient.Close();
                        throw new ConnectionException("连接错误，未找到本地端口");
                    }

                    udpClient.Client.ReceiveTimeout = 5000;
                    endPoint = new IPEndPoint(IPAddress.Parse(IpAddress), _portOut);
                    data = udpClient.Receive(ref endPoint);
                }
                else
                {
                    if (_stream != null)
                    {
                        _stream.Write(data, 0, data.Length - 2);

                        if (SendDataChanged != null)
                        {
                            SendData = new byte[data.Length - 2];
                            Array.Copy(data, 0, SendData, 0, data.Length - 2);
                            SendDataChanged(this);
                        }

                        data = new byte[2100];
                        var numberOfBytes = _stream.Read(data, 0, data.Length);
                        if (ReceiveDataChanged != null)
                        {
                            ReceiveData = new byte[numberOfBytes];
                            Array.Copy(data, 0, ReceiveData, 0, numberOfBytes);
                            ReceiveDataChanged(this);
                        }
                    }
                    else
                    {
                        throw new ConnectionException("连接错误，stream 为空");
                    }
                }
            }
            else
            {
                throw new ConnectionException("连接错误，未找到连接");
            }

            CheckErrorCode(data[8]);

            if (_serialPort != null)
            {
                _crc = BitConverter.GetBytes(CalculateCrc(data, 6, 6));
                if ((_crc[0] != data[12] | _crc[1] != data[13]) & _dataReceived)
                {
                    if (NumberOfRetries <= _countRetries)
                    {
                        _countRetries = 0;
                        throw new CRCCheckFailedException("响应 CRC 校验失败");
                    }

                    //失败重试
                    _countRetries++;
                    WriteSingleCoil(startingAddress, value);
                }
                else if (!_dataReceived)
                {
                    if (NumberOfRetries <= _countRetries)
                    {
                        _countRetries = 0;
                        throw new TimeoutException("Modbus 从站无响应");
                    }

                    //失败重试
                    _countRetries++;
                    WriteSingleCoil(startingAddress, value);
                }
            }
        }


        /// <summary>
        /// 写单个保持寄存器 (FC6).
        /// </summary>
        /// <param name="startingAddress">起始地址</param>
        /// <param name="value">值</param>
        public void WriteSingleRegister(int startingAddress, int value)
        {
            _transactionIdentifierInternal++;
            if (_serialPort is { IsOpen: false })
            {
                throw new SerialPortNotOpenedException("串行端口未打开");
            }

            if (_tcpClient == null & !UdpFlag & _serialPort == null)
            {
                throw new ConnectionException("连接错误");
            }

            _transactionIdentifier = BitConverter.GetBytes(_transactionIdentifierInternal);
            _protocolIdentifier = BitConverter.GetBytes(0x0000);
            _length = BitConverter.GetBytes(0x0006);
            _functionCode = 0x06;
            _startingAddress = BitConverter.GetBytes(startingAddress);
            var registerValue = BitConverter.GetBytes(value);

            byte[] data =
            [
                _transactionIdentifier[1],
                _transactionIdentifier[0],
                _protocolIdentifier[1],
                _protocolIdentifier[0],
                _length[1],
                _length[0],
                UnitIdentifier,
                _functionCode,
                _startingAddress[1],
                _startingAddress[0],
                registerValue[1],
                registerValue[0],
                _crc[0],
                _crc[1]
            ];
            _crc = BitConverter.GetBytes(CalculateCrc(data, 6, 6));
            data[12] = _crc[0];
            data[13] = _crc[1];
            if (_serialPort != null)
            {
                _dataReceived = false;
                _bytesToRead = 8;
                //RTU模式没有事务标识符，协议标识符和长度，共 6 字节
                _serialPort.Write(data, 6, 8);
                if (SendDataChanged != null)
                {
                    SendData = new byte[8];
                    Array.Copy(data, 6, SendData, 0, 8);
                    SendDataChanged(this);
                }

                data = new byte[2100];
                _readBuffer = new byte[256];
                var dateTimeSend = DateTime.Now;
                byte receivedUnitIdentifier = 0xFF; //接收到的从站地址

                // 10000L = 1 ms，接收数据
                while (receivedUnitIdentifier != UnitIdentifier & !(DateTime.Now.Ticks - dateTimeSend.Ticks >
                                                                    TimeSpan.TicksPerMillisecond *
                                                                    ConnectionTimeout))
                {
                    //等待数据接收完成
                    while (_dataReceived == false & !(DateTime.Now.Ticks - dateTimeSend.Ticks >
                                                      TimeSpan.TicksPerMillisecond * ConnectionTimeout))
                        Thread.Sleep(1);

                    data = new byte[2100];
                    Array.Copy(_readBuffer, 0, data, 6, _readBuffer.Length);
                    receivedUnitIdentifier = data[6];
                }

                //检查从站地址是否匹配
                if (receivedUnitIdentifier != UnitIdentifier)
                    data = new byte[2100];
                else
                    _countRetries = 0;
            }
            else if (_tcpClient != null && _tcpClient.Client.Connected | UdpFlag)
            {
                if (UdpFlag)
                {
                    var udpClient = new UdpClient();
                    var endPoint = new IPEndPoint(IPAddress.Parse(IpAddress), Port);
                    udpClient.Send(data, data.Length - 2, endPoint);

                    if (udpClient.Client.LocalEndPoint is IPEndPoint point)
                        _portOut = point.Port;
                    else
                    {
                        udpClient.Close();
                        throw new ConnectionException("连接错误，未找到本地端口");
                    }

                    udpClient.Client.ReceiveTimeout = 5000;
                    endPoint = new IPEndPoint(IPAddress.Parse(IpAddress), _portOut);
                    data = udpClient.Receive(ref endPoint);
                }
                else
                {
                    if (_stream != null)
                    {
                        _stream.Write(data, 0, data.Length - 2);

                        if (SendDataChanged != null)
                        {
                            SendData = new byte[data.Length - 2];
                            Array.Copy(data, 0, SendData, 0, data.Length - 2);
                            SendDataChanged(this);
                        }

                        data = new byte[2100];
                        var numberOfBytes = _stream.Read(data, 0, data.Length);
                        if (ReceiveDataChanged != null)
                        {
                            ReceiveData = new byte[numberOfBytes];
                            Array.Copy(data, 0, ReceiveData, 0, numberOfBytes);
                            ReceiveDataChanged(this);
                        }
                    }
                    else
                    {
                        throw new ConnectionException("连接错误，stream 为空");
                    }
                }
            }
            else
            {
                throw new ConnectionException("连接错误，未找到连接");
            }

            CheckErrorCode(data[8]);

            if (_serialPort == null) return;
            _crc = BitConverter.GetBytes(CalculateCrc(data, 6, 6));
            if ((_crc[0] != data[12] | _crc[1] != data[13]) & _dataReceived)
            {
                if (NumberOfRetries <= _countRetries)
                {
                    _countRetries = 0;
                    throw new CRCCheckFailedException("响应 CRC 校验失败");
                }

                //重试
                _countRetries++;
                WriteSingleRegister(startingAddress, value);
            }
            else if (!_dataReceived)
            {
                if (NumberOfRetries <= _countRetries)
                {
                    _countRetries = 0;
                    throw new TimeoutException("Modbus 从站无响应");
                }

                //重试
                _countRetries++;
                WriteSingleRegister(startingAddress, value);
            }
        }

        /// <summary>
        /// 写多个线圈 (FC15).
        /// </summary>
        /// <param name="startingAddress">起始地址</param>
        /// <param name="values">值</param>
        public void WriteMultipleCoils(int startingAddress, bool[] values)
        {
            _transactionIdentifierInternal++;
            var byteCount = (byte)((values.Length % 8 != 0 ? values.Length / 8 + 1 : (values.Length / 8)));
            var quantityOfOutputs = BitConverter.GetBytes(values.Length);
            byte singleCoilValue = 0;
            if (_serialPort is { IsOpen: false })
            {
                throw new SerialPortNotOpenedException("串行端口未打开");
            }

            if (_tcpClient == null & !UdpFlag & _serialPort == null)
            {
                throw new ConnectionException("连接错误");
            }

            _transactionIdentifier = BitConverter.GetBytes(_transactionIdentifierInternal);
            _protocolIdentifier = BitConverter.GetBytes(0x0000);
            _length = BitConverter.GetBytes(7 + (byteCount));
            _functionCode = 0x0F;
            _startingAddress = BitConverter.GetBytes(startingAddress);


            var data = new byte[14 + 2 + (values.Length % 8 != 0 ? values.Length / 8 : (values.Length / 8) - 1)];
            data[0] = _transactionIdentifier[1];
            data[1] = _transactionIdentifier[0];
            data[2] = _protocolIdentifier[1];
            data[3] = _protocolIdentifier[0];
            data[4] = _length[1];
            data[5] = _length[0];
            data[6] = UnitIdentifier;
            data[7] = _functionCode;
            data[8] = _startingAddress[1];
            data[9] = _startingAddress[0];
            data[10] = quantityOfOutputs[1];
            data[11] = quantityOfOutputs[0];
            data[12] = byteCount;
            for (var i = 0; i < values.Length; i++)
            {
                if ((i % 8) == 0)
                    singleCoilValue = 0;
                var coilValue = values[i] ? (byte)1 : (byte)0;

                singleCoilValue = (byte)(coilValue << (i % 8) | singleCoilValue);

                data[13 + (i / 8)] = singleCoilValue;
            }

            _crc = BitConverter.GetBytes(CalculateCrc(data, (ushort)(data.Length - 8), 6));
            data[^2] = _crc[0];
            data[^1] = _crc[1];
            if (_serialPort != null)
            {
                _dataReceived = false;
                _bytesToRead = 8;
                //RTU模式没有事务标识符，协议标识符和长度，共 6 字节
                _serialPort.Write(data, 6, data.Length - 6);

                if (SendDataChanged != null)
                {
                    SendData = new byte[data.Length - 6];
                    Array.Copy(data, 6, SendData, 0, data.Length - 6);
                    SendDataChanged(this);
                }

                data = new byte[2100];
                _readBuffer = new byte[256];
                var dateTimeSend = DateTime.Now;
                byte receivedUnitIdentifier = 0xFF; //接收到的从站地址

                // 10000L = 1 ms，接收数据
                while (receivedUnitIdentifier != UnitIdentifier & !((DateTime.Now.Ticks - dateTimeSend.Ticks) >
                                                                    TimeSpan.TicksPerMillisecond *
                                                                    ConnectionTimeout))
                {
                    //等待数据接收完成
                    while (_dataReceived == false & !((DateTime.Now.Ticks - dateTimeSend.Ticks) >
                                                      TimeSpan.TicksPerMillisecond * ConnectionTimeout))
                        Thread.Sleep(1);
                    data = new byte[2100];
                    Array.Copy(_readBuffer, 0, data, 6, _readBuffer.Length);
                    receivedUnitIdentifier = data[6];
                }

                //检查从站地址是否匹配
                if (receivedUnitIdentifier != UnitIdentifier)
                    data = new byte[2100];
                else
                    _countRetries = 0;
            }
            else if (_tcpClient != null && _tcpClient.Client.Connected | UdpFlag)
            {
                if (UdpFlag)
                {
                    var udpClient = new UdpClient();
                    var endPoint = new IPEndPoint(IPAddress.Parse(IpAddress), Port);
                    udpClient.Send(data, data.Length - 2, endPoint);

                    if (udpClient.Client.LocalEndPoint is IPEndPoint point)
                        _portOut = point.Port;
                    else
                    {
                        udpClient.Close();
                        throw new ConnectionException("连接错误，未找到本地端口");
                    }

                    udpClient.Client.ReceiveTimeout = 5000;
                    endPoint = new IPEndPoint(IPAddress.Parse(IpAddress), _portOut);
                    data = udpClient.Receive(ref endPoint);
                }
                else
                {
                    if (_stream != null)
                    {
                        _stream.Write(data, 0, data.Length - 2);

                        if (SendDataChanged != null)
                        {
                            SendData = new byte[data.Length - 2];
                            Array.Copy(data, 0, SendData, 0, data.Length - 2);
                            SendDataChanged(this);
                        }

                        data = new byte[2100];
                        var numberOfBytes = _stream.Read(data, 0, data.Length);
                        if (ReceiveDataChanged != null)
                        {
                            ReceiveData = new byte[numberOfBytes];
                            Array.Copy(data, 0, ReceiveData, 0, numberOfBytes);
                            ReceiveDataChanged(this);
                        }
                    }
                    else
                    {
                        throw new ConnectionException("连接错误，stream 为空");
                    }
                }
            }
            else
            {
                throw new ConnectionException("连接错误，未找到连接");
            }

            CheckErrorCode(data[8]);

            if (_serialPort == null) return;

            _crc = BitConverter.GetBytes(CalculateCrc(data, 6, 6));
            if ((_crc[0] != data[12] | _crc[1] != data[13]) & _dataReceived)
            {
                if (NumberOfRetries <= _countRetries)
                {
                    _countRetries = 0;
                    throw new CRCCheckFailedException("响应 CRC 校验失败");
                }

                _countRetries++;
                WriteMultipleCoils(startingAddress, values);
            }
            else if (!_dataReceived)
            {
                if (NumberOfRetries <= _countRetries)
                {
                    _countRetries = 0;
                    throw new TimeoutException("Modbus 从站无响应");
                }

                _countRetries++;
                WriteMultipleCoils(startingAddress, values);
            }
        }

        /// <summary>
        /// 写多个保持寄存器 (FC16).
        /// </summary>
        /// <param name="startingAddress">起始地址</param>
        /// <param name="values">值</param>
        public void WriteMultipleRegisters(int startingAddress, int[] values)
        {
            _transactionIdentifierInternal++;
            var byteCount = (byte)(values.Length * 2);
            var quantityOfOutputs = BitConverter.GetBytes(values.Length);
            if (_serialPort is { IsOpen: false })
            {
                throw new SerialPortNotOpenedException("串行端口未打开");
            }

            if (_tcpClient == null & !UdpFlag & _serialPort == null)
            {
                throw new ConnectionException("连接错误");
            }

            _transactionIdentifier = BitConverter.GetBytes(_transactionIdentifierInternal);
            _protocolIdentifier = BitConverter.GetBytes(0x0000);
            _length = BitConverter.GetBytes(7 + values.Length * 2);
            _functionCode = 0x10;
            _startingAddress = BitConverter.GetBytes(startingAddress);

            var data = new byte[13 + 2 + values.Length * 2];
            data[0] = _transactionIdentifier[1];
            data[1] = _transactionIdentifier[0];
            data[2] = _protocolIdentifier[1];
            data[3] = _protocolIdentifier[0];
            data[4] = _length[1];
            data[5] = _length[0];
            data[6] = UnitIdentifier;
            data[7] = _functionCode;
            data[8] = _startingAddress[1];
            data[9] = _startingAddress[0];
            data[10] = quantityOfOutputs[1];
            data[11] = quantityOfOutputs[0];
            data[12] = byteCount;
            for (var i = 0; i < values.Length; i++)
            {
                var singleRegisterValue = BitConverter.GetBytes(values[i]);
                data[13 + i * 2] = singleRegisterValue[1];
                data[14 + i * 2] = singleRegisterValue[0];
            }

            _crc = BitConverter.GetBytes(CalculateCrc(data, (ushort)(data.Length - 8), 6));
            data[^2] = _crc[0];
            data[^1] = _crc[1];
            if (_serialPort != null)
            {
                _dataReceived = false;
                _bytesToRead = 8;
                _serialPort.Write(data, 6, data.Length - 6);

                if (SendDataChanged != null)
                {
                    SendData = new byte[data.Length - 6];
                    Array.Copy(data, 6, SendData, 0, data.Length - 6);
                    SendDataChanged(this);
                }

                data = new byte[2100];
                _readBuffer = new byte[256];
                var dateTimeSend = DateTime.Now;
                byte receivedUnitIdentifier = 0xFF; //接收到的从站地址

                // 10000L = 1 ms，接收数据
                while (receivedUnitIdentifier != UnitIdentifier & !((DateTime.Now.Ticks - dateTimeSend.Ticks) >
                                                                    TimeSpan.TicksPerMillisecond *
                                                                    ConnectionTimeout))
                {
                    //等待数据接收完成
                    while (_dataReceived == false & !((DateTime.Now.Ticks - dateTimeSend.Ticks) >
                                                      TimeSpan.TicksPerMillisecond * ConnectionTimeout))
                        Thread.Sleep(1);
                    data = new byte[2100];
                    Array.Copy(_readBuffer, 0, data, 6, _readBuffer.Length);
                    receivedUnitIdentifier = data[6];
                }

                //检查从站地址是否匹配
                if (receivedUnitIdentifier != UnitIdentifier)
                    data = new byte[2100];
                else
                    _countRetries = 0;
            }
            else if (_tcpClient != null && _tcpClient.Client.Connected | UdpFlag)
            {
                if (UdpFlag)
                {
                    var udpClient = new UdpClient();
                    var endPoint = new IPEndPoint(IPAddress.Parse(IpAddress), Port);
                    udpClient.Send(data, data.Length - 2, endPoint);

                    if (udpClient.Client.LocalEndPoint is IPEndPoint point)
                        _portOut = point.Port;
                    else
                    {
                        udpClient.Close();
                        throw new ConnectionException("连接错误，未找到本地端口");
                    }

                    udpClient.Client.ReceiveTimeout = 5000;
                    endPoint = new IPEndPoint(IPAddress.Parse(IpAddress), _portOut);
                    data = udpClient.Receive(ref endPoint);
                }
                else
                {
                    if (_stream != null)
                    {
                        _stream.Write(data, 0, data.Length - 2);

                        if (SendDataChanged != null)
                        {
                            SendData = new byte[data.Length - 2];
                            Array.Copy(data, 0, SendData, 0, data.Length - 2);
                            SendDataChanged(this);
                        }

                        data = new byte[2100];
                        var numberOfBytes = _stream.Read(data, 0, data.Length);
                        if (ReceiveDataChanged != null)
                        {
                            ReceiveData = new byte[numberOfBytes];
                            Array.Copy(data, 0, ReceiveData, 0, numberOfBytes);
                            ReceiveDataChanged(this);
                        }
                    }
                    else
                    {
                        throw new ConnectionException("连接错误，stream 为空");
                    }
                }
            }
            else
            {
                throw new ConnectionException("连接错误，未找到连接");
            }

            CheckErrorCode(data[8]);

            if (_serialPort == null)
                return;

            _crc = BitConverter.GetBytes(CalculateCrc(data, 6, 6));
            if ((_crc[0] != data[12] | _crc[1] != data[13]) & _dataReceived)
            {
                if (NumberOfRetries <= _countRetries)
                {
                    _countRetries = 0;
                    throw new CRCCheckFailedException("响应 CRC 校验失败");
                }

                _countRetries++;
                WriteMultipleRegisters(startingAddress, values);
            }
            else if (!_dataReceived)
            {
                if (NumberOfRetries <= _countRetries)
                {
                    _countRetries = 0;
                    throw new TimeoutException("Modbus 从站无响应");
                }

                _countRetries++;
                WriteMultipleRegisters(startingAddress, values);
            }
        }

        /// <summary>
        /// 读、写多个寄存器 (FC23).
        /// </summary>
        /// <param name="startingAddressRead">读取的起始地址</param>
        /// <param name="quantityRead">读取的寄存器个数</param>
        /// <param name="startingAddressWrite">写入的起始地址</param>
        /// <param name="values">写入的值</param>
        /// <returns>读取的结果</returns>
        public int[] ReadWriteMultipleRegisters(int startingAddressRead, int quantityRead, int startingAddressWrite,
            int[] values)
        {
            _transactionIdentifierInternal++;
            if (_serialPort is { IsOpen: false })
            {
                throw new SerialPortNotOpenedException("串行端口未打开");
            }

            if (_tcpClient == null & !UdpFlag & _serialPort == null)
            {
                throw new ConnectionException("连接错误");
            }

            if (startingAddressRead > 65535 | quantityRead > 125 | startingAddressWrite > 65535 | values.Length > 121)
            {
                throw new ArgumentException("起始地址必须为 0 - 65535；查询的数量必须为 1 - 125；写入的数量必须为 1 - 121");
            }

            _transactionIdentifier = BitConverter.GetBytes(_transactionIdentifierInternal);
            _protocolIdentifier = BitConverter.GetBytes(0x0000);
            _length = BitConverter.GetBytes(11 + values.Length * 2);
            _functionCode = 0x17;
            var startingAddressReadLocal = BitConverter.GetBytes(startingAddressRead);
            var quantityReadLocal = BitConverter.GetBytes(quantityRead);
            var startingAddressWriteLocal = BitConverter.GetBytes(startingAddressWrite);
            var quantityWriteLocal = BitConverter.GetBytes(values.Length);
            var writeByteCountLocal = Convert.ToByte(values.Length * 2);
            var data = new byte[17 + 2 + values.Length * 2];
            data[0] = _transactionIdentifier[1];
            data[1] = _transactionIdentifier[0];
            data[2] = _protocolIdentifier[1];
            data[3] = _protocolIdentifier[0];
            data[4] = _length[1];
            data[5] = _length[0];
            data[6] = UnitIdentifier;
            data[7] = _functionCode;
            data[8] = startingAddressReadLocal[1];
            data[9] = startingAddressReadLocal[0];
            data[10] = quantityReadLocal[1];
            data[11] = quantityReadLocal[0];
            data[12] = startingAddressWriteLocal[1];
            data[13] = startingAddressWriteLocal[0];
            data[14] = quantityWriteLocal[1];
            data[15] = quantityWriteLocal[0];
            data[16] = writeByteCountLocal;

            for (var i = 0; i < values.Length; i++)
            {
                var singleRegisterValue = BitConverter.GetBytes(values[i]);
                data[17 + i * 2] = singleRegisterValue[1];
                data[18 + i * 2] = singleRegisterValue[0];
            }

            _crc = BitConverter.GetBytes(CalculateCrc(data, (ushort)(data.Length - 8), 6));
            data[^2] = _crc[0];
            data[^1] = _crc[1];
            if (_serialPort != null)
            {
                _dataReceived = false;
                _bytesToRead = 5 + 2 * quantityRead;
                _serialPort.Write(data, 6, data.Length - 6);

                if (SendDataChanged != null)
                {
                    SendData = new byte[data.Length - 6];
                    Array.Copy(data, 6, SendData, 0, data.Length - 6);
                    SendDataChanged(this);
                }

                data = new byte[2100];
                _readBuffer = new byte[256];
                var dateTimeSend = DateTime.Now;
                byte receivedUnitIdentifier = 0xFF; //接收到的从站地址

                // 10000L = 1 ms，接收数据
                while (receivedUnitIdentifier != UnitIdentifier & !((DateTime.Now.Ticks - dateTimeSend.Ticks) >
                                                                    TimeSpan.TicksPerMillisecond *
                                                                    ConnectionTimeout))
                {
                    //等待数据接收完成
                    while (_dataReceived == false & !((DateTime.Now.Ticks - dateTimeSend.Ticks) >
                                                      TimeSpan.TicksPerMillisecond * ConnectionTimeout))
                        Thread.Sleep(1);
                    data = new byte[2100];
                    Array.Copy(_readBuffer, 0, data, 6, _readBuffer.Length);
                    receivedUnitIdentifier = data[6];
                }

                //检查从站地址是否匹配
                if (receivedUnitIdentifier != UnitIdentifier)
                    data = new byte[2100];
                else
                    _countRetries = 0;
            }
            else if (_tcpClient != null && _tcpClient.Client.Connected | UdpFlag)
            {
                if (UdpFlag)
                {
                    var udpClient = new UdpClient();
                    var endPoint = new IPEndPoint(IPAddress.Parse(IpAddress), Port);
                    udpClient.Send(data, data.Length - 2, endPoint);

                    if (udpClient.Client.LocalEndPoint is IPEndPoint point)
                        _portOut = point.Port;
                    else
                    {
                        udpClient.Close();
                        throw new ConnectionException("连接错误，未找到本地端口");
                    }

                    udpClient.Client.ReceiveTimeout = 5000;
                    endPoint = new IPEndPoint(IPAddress.Parse(IpAddress), _portOut);
                    data = udpClient.Receive(ref endPoint);
                }
                else
                {
                    if (_stream != null)
                    {
                        _stream.Write(data, 0, data.Length - 2);

                        if (SendDataChanged != null)
                        {
                            SendData = new byte[data.Length - 2];
                            Array.Copy(data, 0, SendData, 0, data.Length - 2);
                            SendDataChanged(this);
                        }

                        data = new byte[2100];
                        var numberOfBytes = _stream.Read(data, 0, data.Length);
                        if (ReceiveDataChanged != null)
                        {
                            ReceiveData = new byte[numberOfBytes];
                            Array.Copy(data, 0, ReceiveData, 0, numberOfBytes);
                            ReceiveDataChanged(this);
                        }
                    }
                    else
                    {
                        throw new ConnectionException("连接错误，stream 为空");
                    }
                }
            }
            else
            {
                throw new ConnectionException("连接错误，未找到连接");
            }

            CheckErrorCode(data[8]);

            var response = new int[quantityRead];
            for (var i = 0; i < quantityRead; i++)
            {
                var highByte = data[9 + i * 2];
                var lowByte = data[9 + i * 2 + 1];

                data[9 + i * 2] = lowByte;
                data[9 + i * 2 + 1] = highByte;

                response[i] = BitConverter.ToInt16(data, (9 + i * 2));
            }

            return response;
        }

        /// <summary>
        /// 关闭与主站的连接。
        /// </summary>
        public void Disconnect()
        {
            if (_serialPort != null)
            {
                if (_serialPort.IsOpen & !_receiveActive)
                    _serialPort.Close();
                ConnectedChanged?.Invoke(this);
                return;
            }

            _stream?.Close();
            _tcpClient?.Close();
            ConnectedChanged?.Invoke(this);
        }

        /// <summary>
        /// 析构函数 - 关闭与主设备的连接。
        /// </summary>
        ~ModbusClient()
        {
            if (_serialPort != null)
            {
                if (_serialPort.IsOpen)
                    _serialPort.Close();
                return;
            }

            if (_tcpClient != null & !UdpFlag)
            {
                _stream?.Close();
                _tcpClient?.Close();
            }
        }

        /// <summary>
        /// 如果客户端已连接服务器，则返回 "TRUE"；如果未连接，则返回 "FALSE"。对于 Modbus RTU，如果 COM 端口已打开，则返回 "TRUE"。
        /// </summary>
        public bool Connected
        {
            get
            {
                if (_serialPort != null)
                    return _serialPort.IsOpen;
                if (UdpFlag & _tcpClient != null)
                    return true;
                return _tcpClient is { Connected: true };
            }
        }

        public bool Available(int timeout)
        {
            // Ping
            var pingSender = new Ping();
            var address = IPAddress.Parse(IpAddress);

            // 创建一个包含 32 字节待传输数据的缓冲区。
            // ReSharper disable once StringLiteralTypo
            const string data = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
            var buffer = Encoding.ASCII.GetBytes(data);
            // 等待回复，有超时设置。
            var reply = pingSender.Send(address, timeout, buffer);
            return reply.Status == IPStatus.Success;
        }

        /// <summary>
        /// 获取或设置服务器的 IP 地址。
        /// </summary>
        // ReSharper disable once MemberCanBePrivate.Global
        public string IpAddress { get; set; } = "127.0.0.1";

        /// <summary>
        /// 获取或设置可连接 Modbus-TCP 服务器的端口（标准协议为 502）。
        /// </summary>
        // ReSharper disable once MemberCanBePrivate.Global
        public int Port { get; set; } = 502;

        /// <summary>
        /// 获取或设置 UDP 标志，以激活 Modbus UDP。
        /// </summary>
        // ReSharper disable once MemberCanBePrivate.Global
        public bool UdpFlag { get; set; } = false;

        /// <summary>
        /// 获取或设置串行连接时的单位标识符（默认 = 0）
        /// </summary>
        // ReSharper disable once MemberCanBePrivate.Global
        public byte UnitIdentifier { get; set; } = 0x01;


        /// <summary>
        /// 获取或设置串行连接的波特率（默认值 = 9600）
        /// </summary>
        // ReSharper disable once MemberCanBePrivate.Global
        public int Baud { get; set; } = 9600;

        /// <summary>
        /// 获取或设置串行连接时的奇偶校验值
        /// </summary>
        // ReSharper disable once MemberCanBePrivate.Global
        public Parity Parity
        {
            get => _serialPort != null ? _parity : Parity.Even;
            set
            {
                if (_serialPort != null)
                    _parity = value;
            }
        }


        /// <summary>
        /// 获取或设置串行连接时的停止位数
        /// </summary>
        // ReSharper disable once MemberCanBePrivate.Global
        public StopBits StopBits
        {
            get => _serialPort != null ? _stopBits : StopBits.One;
            set
            {
                if (_serialPort != null)
                    _stopBits = value;
            }
        }

        /// <summary>
        /// 获取或设置 ModbusTCP 连接的连接超时
        /// </summary>
        // ReSharper disable once MemberCanBePrivate.Global
        public int ConnectionTimeout { get; set; } = 1000;

        /// <summary>
        /// 获取或设置串行端口
        /// </summary>
        // ReSharper disable once MemberCanBePrivate.Global
        public string? SerialPort
        {
            get => _serialPort?.PortName;
            set
            {
                if (value == null)
                {
                    _serialPort = null;
                    return;
                }

                _serialPort?.Close();
                _serialPort = new SerialPort();
                _serialPort.PortName = value;
                _serialPort.BaudRate = Baud;
                _serialPort.Parity = _parity;
                _serialPort.StopBits = _stopBits;
                _serialPort.WriteTimeout = 10000;
                _serialPort.ReadTimeout = ConnectionTimeout;
                _serialPort.DataReceived += DataReceivedHandler;
            }
        }
    }
}