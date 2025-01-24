namespace Toxy;

public interface IPrettyTable
{
    string GetString();
    string GetString(int startCol, int endCol);
    string GetString(string[] fieldRange);
}
