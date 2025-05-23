namespace Excel2Json;

public class ProcessError(string message, int row, int column, bool isWarning)
{
    public string Message { get; } = message;
    public int Row { get; } = row;
    public int Column { get; } = column;
    public bool IsWarning { get; } = isWarning;
} 