using System;

public class PageNumMismatchException : Exception
{
    public PageNumMismatchException()
    {
    }

    public PageNumMismatchException(string message)
        : base(message)
    {
    }

    public PageNumMismatchException(string message, Exception inner)
        : base(message, inner)
    {
    }
}