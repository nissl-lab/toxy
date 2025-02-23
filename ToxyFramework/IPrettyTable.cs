namespace Toxy;

public interface IPrettyTable
{
    string Print();
    string Print(int startCol, int endCol);
    string Print(string[] fieldRange);
}
