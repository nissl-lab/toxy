namespace Toxy;

public interface IPrettyTable
{
    void Print();
    void Print(int startCol, int endCol);
    void Print(string[] fieldRange);
}
