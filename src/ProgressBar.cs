namespace ExcelCsvMerger;

/// <summary>
///  For displaying progress to the user.
/// </summary>
public class ProgressBar
{
    /// <summary>
    /// Initializes a new <see cref="ProgressBar"/> object, with
    /// title and size specified.
    /// </summary>
    /// <param name="title">Text to display before progress bar.</param>
    /// <param name="size">Total count of collection (used to determine progress percentage).</param>
    public ProgressBar(string title, int size)
    {
        this.title = title;
        this.size = size;
    }
    
    private readonly string title;
    private readonly int size;
    private int count;

    /// <summary>
    /// Increments the progress bar by 1, rendering a new percentage.
    /// </summary>
    public void IncrementProgressBar()
    {
        const int progressBarWidth = 80;
        var progress = (double)++count / size;
        var completedWidth = (int)(progress * progressBarWidth);

        Console.Write($"{title} [");
        for (var i = 0; i < progressBarWidth; i++)
        {
            Console.Write(i < completedWidth ? "=" : " ");
        }
        
        Console.Write($"] ({count}/{size})\r"); // \r moves the cursor to the beginning of the line

        if (count == size)
        {
            Console.WriteLine("");
        }
    }
}