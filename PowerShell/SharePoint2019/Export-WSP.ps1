$dirName = Get-Location

foreach ($solution in Get-SPSolution)
{
    $id = $Solution.SolutionID
    $title = $Solution.Name
    $filename = $Solution.SolutionFile.Name
    $solution.SolutionFile.SaveAs("$dirName\$filename")
}
