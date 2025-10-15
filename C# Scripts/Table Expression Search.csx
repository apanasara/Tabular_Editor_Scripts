using System.Diagnostics.CodeAnalysis;

string searchKeyword = "eplica";
string searchInto = "";
string op = "";

// Iterate over all tables in the model
foreach (var table in Model.Tables)
{
    //op = $"{op}"+ Environment.NewLine +$"Table: {table.Name} ({table.Partitions.Count}) - {table.Partitions[0].Quer}";
    

    //------------Getting Power query-------------
    searchInto = "";
    if (table != null && table.Partitions.Count > 1)
    {
        searchInto = $"{table.SourceExpression}";
    }
    else if (table != null && table.Partitions.Count == 1)
    {
        searchInto = $"{table.Partitions[0].Query}";
    }

    // ------------ Searching Keyword ----------------
    if ( searchInto.Contains(searchKeyword) )
    {
        op = $"{op}"+ Environment.NewLine +$"{table.Name}";
    }


}

Output(op);