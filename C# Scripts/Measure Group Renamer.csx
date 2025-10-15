string sourceKeyword = "12";
string creatingKeyword = "36";

// Iterate over selected measures in the model
foreach (var measure in Selected.Measures)
{
    // Check if the measure name contains the keyword
    if (measure.Name.Contains(sourceKeyword) || measure.Expression.Contains(sourceKeyword))
    {
        measure.Name = measure.Name.Replace(sourceKeyword, creatingKeyword);
        measure.Name = measure.Name.Replace(" 1", "");
        
        measure.Expression = measure.Expression.Replace(sourceKeyword, creatingKeyword);
        measure.DisplayFolder = measure.DisplayFolder.Replace(sourceKeyword, creatingKeyword);
    }
}
Info("Finished");