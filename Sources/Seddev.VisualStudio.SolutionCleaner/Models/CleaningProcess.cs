namespace Seddev.VisualStudio.SolutionCleaner.Models
{
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    internal class CleaningProcess
    {
        public CleaningProcess(string solutionPath)
        {
            SolutionProjects = new List<Project>();
            
            var projectPaths = Directory.GetFiles(solutionPath, "*.csproj", SearchOption.AllDirectories);

            for (var i = 0; i < projectPaths.Count(); i++)
            {
                var fileInfo = new System.IO.FileInfo(projectPaths[i]);

                SolutionProjects.Add(new Project
                {
                    Id = i + 1, 
                    Path = fileInfo.DirectoryName,
                    Name = fileInfo.Name
                });
            }
        }

        public List<Project> SolutionProjects { get; set; }

        public int SuccessCount { get; set; }
        
        public int FailedCount { get; set; }        

    }
}
