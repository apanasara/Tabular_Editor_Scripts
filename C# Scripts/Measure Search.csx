string searchKeyword = "TREATAS(";
string op = "";

// Iterate over selected measures in the model
foreach (var measure in Selected.Measures)
{
    // Check if the measure name contains the keyword
    if ( measure.Expression.Contains(searchKeyword))
    {        
        //measure.Highlight();
        op = $"{op}"+ Environment.NewLine +$"{measure.DisplayFolder} >> {measure.Name}";
    }
}
Output(op);