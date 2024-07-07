using System.Runtime.Serialization;

namespace ModbusTCP_RTU.Exceptions;

/// <summary>
///     如果串行端口未打开，将抛出异常
/// </summary>
public class SerialPortNotOpenedException : ModbusException
{
    public SerialPortNotOpenedException()
    {
    }

    public SerialPortNotOpenedException(string message)
        : base(message)
    {
    }

    public SerialPortNotOpenedException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    protected SerialPortNotOpenedException(SerializationInfo info, StreamingContext context)
        : base(info, context)
    {
    }
}

/// <summary>
///     如果与 Modbus 设备连接失败，将抛出异常
/// </summary>
public class ConnectionException : ModbusException
{
    public ConnectionException()
    {
    }

    public ConnectionException(string message)
        : base(message)
    {
    }

    public ConnectionException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    protected ConnectionException(SerializationInfo info, StreamingContext context)
        : base(info, context)
    {
    }
}

/// <summary>
///     如果 Modbus 服务器返回错误代码 "不支持功能代码"，则抛出异常
/// </summary>
public class FunctionCodeNotSupportedException : ModbusException
{
    public FunctionCodeNotSupportedException()
    {
    }

    public FunctionCodeNotSupportedException(string message)
        : base(message)
    {
    }

    public FunctionCodeNotSupportedException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    protected FunctionCodeNotSupportedException(SerializationInfo info, StreamingContext context)
        : base(info, context)
    {
    }
}

/// <summary>
///     如果 Modbus 服务器返回错误代码 "数量无效"，则抛出异常
/// </summary>
public class QuantityInvalidException : ModbusException
{
    public QuantityInvalidException()
    {
    }

    public QuantityInvalidException(string message)
        : base(message)
    {
    }

    public QuantityInvalidException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    protected QuantityInvalidException(SerializationInfo info, StreamingContext context)
        : base(info, context)
    {
    }
}

/// <summary>
///     如果 Modbus 服务器返回错误代码 "起始地址和数量无效"，则抛出异常
/// </summary>
public class StartingAddressInvalidException : ModbusException
{
    public StartingAddressInvalidException()
    {
    }

    public StartingAddressInvalidException(string message)
        : base(message)
    {
    }

    public StartingAddressInvalidException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    protected StartingAddressInvalidException(SerializationInfo info, StreamingContext context)
        : base(info, context)
    {
    }
}

/// <summary>
///     如果 Modbus 服务器返回错误代码 "未定义功能代码 (0x04)"，则抛出异常
/// </summary>
public class ModbusException : Exception
{
    public ModbusException()
    {
    }

    public ModbusException(string message)
        : base(message)
    {
    }

    public ModbusException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    protected ModbusException(SerializationInfo info, StreamingContext context)
        : base(info, context)
    {
    }
}

/// <summary>
///     如果 CRC 校验失败，将抛出异常
/// </summary>
public class CRCCheckFailedException : ModbusException
{
    public CRCCheckFailedException()
    {
    }

    public CRCCheckFailedException(string message)
        : base(message)
    {
    }

    public CRCCheckFailedException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    protected CRCCheckFailedException(SerializationInfo info, StreamingContext context)
        : base(info, context)
    {
    }
}