using Microsoft.Extensions.Configuration;

public interface IStrategy
{
    Task ProcessAsync(IEnumerable<string> files);
}
