/*

This Script will replcae string to following paramenters of newly added measure
1. Measure Name
2. Measure's Expressions e.g, Filter string will be altered
3. Measure Description
4. Measure's Display folder will be created by 

*/

string sourceKeyword = "America"; 
string creatingKeyword = "Africa";

// Loop through selected measures
foreach (var measure in Selected.Measures)
{
    // Create a new measure in the same table
    var newMeasure = measure.Table.AddMeasure(
        measure.Name.Replace(sourceKeyword, creatingKeyword),
        measure.Expression.Replace(sourceKeyword, creatingKeyword)
    );

    // Optionally copy formatting and properties
    newMeasure.FormatString = measure.FormatString;
    newMeasure.Description =  measure.Description.Replace(sourceKeyword, creatingKeyword);
    newMeasure.DisplayFolder =  measure.DisplayFolder.Replace(sourceKeyword, creatingKeyword);
}

Info("Finished");
