using APPZ_Lab3;
using APPZ_Lab3.Data_classes;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

List<Vacancy> vacancies = DataBaseConnectionClass.GetVacancies();
List<IUser> users = DataBaseConnectionClass.GetUsersAll();

ExportToWord(vacancies, users);
ExportToExcel(vacancies, users);


void ExportToWord(List<Vacancy> vacancies, List<IUser> users)
{
    Word.Application wordApp = new Word.Application();
    Word.Document doc = wordApp.Documents.Add();

    Word.Paragraph para1 = doc.Paragraphs.Add();
    para1.Range.Text = "Vacancies";
    para1.Range.InsertParagraphAfter();

    Word.Table vacancyTable = doc.Tables.Add(para1.Range, vacancies.Count + 1, 5);
    vacancyTable.Cell(1, 1).Range.Text = "ID";
    vacancyTable.Cell(1, 2).Range.Text = "User ID";
    vacancyTable.Cell(1, 3).Range.Text = "Hourly Rate";
    vacancyTable.Cell(1, 4).Range.Text = "Subject";
    vacancyTable.Cell(1, 5).Range.Text = "Status";

    for (int i = 0; i < vacancies.Count; i++)
    {
        vacancyTable.Cell(i + 2, 1).Range.Text = vacancies[i].Id.ToString();
        vacancyTable.Cell(i + 2, 2).Range.Text = vacancies[i].UserId.ToString();
        vacancyTable.Cell(i + 2, 3).Range.Text = vacancies[i].HourlyRate.ToString();
        vacancyTable.Cell(i + 2, 4).Range.Text = vacancies[i].Subject;
        vacancyTable.Cell(i + 2, 5).Range.Text = vacancies[i].Status;
    }

    Word.Paragraph para2 = doc.Paragraphs.Add();
    para2.Range.Text = "Users";
    para2.Range.InsertParagraphAfter();

    Word.Table userTable = doc.Tables.Add(para2.Range, users.Count + 1, 7);
    userTable.Cell(1, 1).Range.Text = "ID";
    userTable.Cell(1, 2).Range.Text = "Username";
    userTable.Cell(1, 3).Range.Text = "Last Name";
    userTable.Cell(1, 4).Range.Text = "First Name";
    userTable.Cell(1, 5).Range.Text = "Email";
    userTable.Cell(1, 6).Range.Text = "Phone";
    userTable.Cell(1, 7).Range.Text = "Sex";

    for (int i = 0; i < users.Count; i++)
    {
        userTable.Cell(i + 2, 1).Range.Text = users[i].Id.ToString();
        userTable.Cell(i + 2, 2).Range.Text = users[i].Username;
        userTable.Cell(i + 2, 3).Range.Text = users[i].LastName;
        userTable.Cell(i + 2, 4).Range.Text = users[i].FirstName;
        userTable.Cell(i + 2, 5).Range.Text = users[i].Email;
        userTable.Cell(i + 2, 6).Range.Text = users[i].Phone;
        userTable.Cell(i + 2, 7).Range.Text = users[i].sex;
    }

    doc.SaveAs2("Report.docx");
    wordApp.Quit();

    System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);

    GC.Collect();
    GC.WaitForPendingFinalizers();
}

void ExportToExcel(List<Vacancy> vacancies, List<IUser> users)
{
    Excel.Application xlApp = new Excel.Application();
    Excel.Workbook xlWorkBook = xlApp.Workbooks.Add();
    Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

    xlWorkSheet.Cells[1, 1] = "ID";
    xlWorkSheet.Cells[1, 2] = "User ID";
    xlWorkSheet.Cells[1, 3] = "Hourly Rate";
    xlWorkSheet.Cells[1, 4] = "Subject";
    xlWorkSheet.Cells[1, 5] = "Status";

    for (int i = 0; i < vacancies.Count; i++)
    {
        xlWorkSheet.Cells[i + 2, 1] = vacancies[i].Id.ToString();
        xlWorkSheet.Cells[i + 2, 2] = vacancies[i].UserId.ToString();
        xlWorkSheet.Cells[i + 2, 3] = vacancies[i].HourlyRate.ToString();
        xlWorkSheet.Cells[i + 2, 4] = vacancies[i].Subject;
        xlWorkSheet.Cells[i + 2, 5] = vacancies[i].Status;
    }

    Excel.Worksheet xlWorkSheet2 = (Excel.Worksheet)xlWorkBook.Worksheets.Add();
    xlWorkSheet2.Cells[1, 1] = "ID";
    xlWorkSheet2.Cells[1, 2] = "Username";
    xlWorkSheet2.Cells[1, 3] = "Last Name";
    xlWorkSheet2.Cells[1, 4] = "First Name";
    xlWorkSheet2.Cells[1, 5] = "Email";
    xlWorkSheet2.Cells[1, 6] = "Phone";
    xlWorkSheet2.Cells[1, 7] = "Sex";

    for (int i = 0; i < users.Count; i++)
    {
        xlWorkSheet2.Cells[i + 2, 1] = users[i].Id.ToString();
        xlWorkSheet2.Cells[i + 2, 2] = users[i].Username;
        xlWorkSheet2.Cells[i + 2, 3] = users[i].LastName;
        xlWorkSheet2.Cells[i + 2, 4] = users[i].FirstName;
        xlWorkSheet2.Cells[i + 2, 5] = users[i].Email;
        xlWorkSheet2.Cells[i + 2, 6] = users[i].Phone;
        xlWorkSheet2.Cells[i + 2, 7] = users[i].sex;
    }

    xlWorkBook.SaveAs("Report.xlsx");
    xlWorkBook.Close(true);
    xlApp.Quit();

    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkSheet);
    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

    GC.Collect();
    GC.WaitForPendingFinalizers();
}
