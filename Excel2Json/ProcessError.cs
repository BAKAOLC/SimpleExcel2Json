namespace Excel2Json;

public class ProcessError
{
    public string Message { get; }
    public int Row { get; }
    public int Column { get; }
    public bool IsWarning { get; }

    public ProcessError(string message, int row, int column, bool isWarning)
    {
        Message = message;
        Row = row;
        Column = column;
        IsWarning = isWarning;
    }
} 