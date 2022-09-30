
using IronXL;

Console.WriteLine("Welcome to the Skills Competencey Converter. Please make sure all files are placed in the foler specified by the pather vairbale before the process starts");


string filePath = "C:\\Users\\andrew.pla\\Desktop\\Comps";
//If your going to run this multipletime either change the out path to be different than the in path or maake sure to delete the output file or it will get picked up as input the next time around. 
string outPath = "C:\\Users\\andrew.pla\\Desktop\\Comps";


//Lets get the Files that are in the Directory
string[] filePaths = Directory.GetFiles(filePath, "*.xlsx");

if(filePaths.Count() <= 0)
{
    Console.WriteLine("No files were found in the directory"); 
    return; 
}

string templateFile = filePaths[0];

Console.WriteLine("Using " + templateFile + " as the template file"); 

//Open the Template
WorkBook templateWorkbook = WorkBook.Load(templateFile);
WorkSheet genProfileSheet = templateWorkbook.WorkSheets[1];
WorkSheet cloudDevSheet = templateWorkbook.WorkSheets[2];

//Create new Excel WorkBook document. 
WorkBook outputWorkbook = WorkBook.Create();
outputWorkbook.Metadata.Author = "Columbus SC";
WorkSheet newSheet = outputWorkbook.CreateWorkSheet("Competencey Summary");
//Initialize the workbook rows otherwise we get an out of bounds later
newSheet["A1:BZ60"].Value = "";

//Add All of the Header Values
var allGenProfileLabels = genProfileSheet["B4:B7"] + genProfileSheet["B14:B18"] + genProfileSheet["B25:B29"] + genProfileSheet["B36:B40"] + genProfileSheet["B48:B52"];
for (int i = 0; i < allGenProfileLabels.Count(); i++)
{
    var cellArray = allGenProfileLabels.ToArray();
    string address = cellArray[i].AddressString;
    string value = cellArray[i].StringValue;
    Console.WriteLine("Cell {0} has value '{1}'", address, value);
    //Add data and styles to the new worksheet 
    newSheet.Columns[i + 1].Rows[0].Value = value;
}

var allDevSheetLabels = cloudDevSheet["B4:B9"] + cloudDevSheet["B16:B19"] + cloudDevSheet["B26:B30"] + cloudDevSheet["B37:B40"] + cloudDevSheet["B47:B52"] + cloudDevSheet["B59:B63"] + cloudDevSheet["B70:B73"] + cloudDevSheet["B80:B83"];
for (int i = 0; i < allDevSheetLabels.Count(); i++)
{
    var cellArray = allDevSheetLabels.ToArray();
    string address = cellArray[i].AddressString;
    string value = cellArray[i].StringValue;
    Console.WriteLine("Cell {0} has value '{1}'", address, value);
    //Add data and styles to the new worksheet 
    newSheet.Columns[allGenProfileLabels.Count() + i + 1].Rows[0].Value = value;
}

var acceptedValues = new string[] { "None", "Fundamental", "Novice", "Intermediate", "Advanced", "Expert" };

for (int i = 0; i < filePaths.Count()-1; i++)
{
    string path = filePaths[i];
    Console.WriteLine("Now Processing File: " + path); 
    string perosonName = path.Substring(0, path.IndexOf('-')).Replace(filePath+"\\", "");
    //Open the Template
    WorkBook valuesWorkbook = WorkBook.Load(path);
    genProfileSheet = valuesWorkbook.WorkSheets[1];
    cloudDevSheet = valuesWorkbook.WorkSheets[2];

    var allGenProfileValues = genProfileSheet["D4:D7"] + genProfileSheet["D14:D18"] + genProfileSheet["D25:D29"] + genProfileSheet["D36:D40"] + genProfileSheet["D48:D52"];
    var allDevSheetValues = cloudDevSheet["D4:D9"] + cloudDevSheet["D16:D19"] + cloudDevSheet["D26:D30"] + cloudDevSheet["D37:D40"] + cloudDevSheet["D47:D52"] + cloudDevSheet["D59:D63"] + cloudDevSheet["D70:D73"] + cloudDevSheet["D80:D83"];

    //Write the Name
    newSheet.Columns[0].Rows[i+1].Value = perosonName;

    //Write Values
    for (int z = 0; z < allGenProfileValues.Count(); z++)
    {
        var cellArray = allGenProfileValues.ToArray();
        string address = cellArray[z].AddressString;
        string value = cellArray[z].StringValue;
        if(!acceptedValues.Contains(value))
        {
            value = "Invalid"; 
        }
        Console.WriteLine("Cell {0} has value '{1}'", address, value);
        //Add data and styles to the new worksheet 
        newSheet.Columns[z + 1].Rows[i+1].Value = value;
    }

    for (int z = 0; z < allDevSheetValues.Count(); z++)
    {
        var cellArray = allDevSheetValues.ToArray();
        string address = cellArray[z].AddressString;
        string value = cellArray[z].StringValue;
        Console.WriteLine("Cell {0} has value '{1}'", address, value);
        //Add data and styles to the new worksheet 
        newSheet.Columns[allGenProfileValues.Count() + z + 1].Rows[i+1].Value = value;
    }
    valuesWorkbook.Close();
}

//Save the excel file
outputWorkbook.SaveAs(outPath + "\\Result.xlsx");



