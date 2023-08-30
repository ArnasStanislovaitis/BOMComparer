using BOMComparer.Models;

namespace ComparerAPI
{
    public interface IBomFileService
    {
        ComparedBomFile CompareBomFiles(string sourcePath, string targetPath);
        byte[] WriteComparedBomToExcel(ComparedBomFile comparedBomFile);
    }
}