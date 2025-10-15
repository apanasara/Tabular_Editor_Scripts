string searchKeyword = "[Begin Date when Sign Date Blank]";
string replaceWith = "[Join Date]";

// Iterate over selected measures in the model
foreach (var measure in Selected.Measures)
{
    // Check if the measure name contains the keyword
    if ( measure.Expression.Contains(searchKeyword))
    {        
        //measure.Highlight();
        measure.Expression = measure.Expression.Replace(searchKeyword, replaceWith);
    }
}