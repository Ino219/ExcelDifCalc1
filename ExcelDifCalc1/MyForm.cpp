#include "MyForm.h"

using namespace ExcelDifCalc1;

[STAThreadAttribute]

int main() {
	Application::Run(gcnew MyForm());
	return 0;
}

System::Void ExcelDifCalc1::MyForm::MyForm_Load(System::Object ^ sender, System::EventArgs ^ e)
{
	
	return System::Void();
}

System::Void ExcelDifCalc1::MyForm::file1_DragDrop(System::Object ^ sender, System::Windows::Forms::DragEventArgs ^ e)
{
	//リストとしてファイルパス取得
	array<String^>^ file = (array<String^>^)e->Data->GetData(DataFormats::FileDrop, false);
	//拡張子取得
	String^	extension = System::IO::Path::GetExtension(file[0]);
	//ファイル名取得
	String^ title = System::IO::Path::GetFileName(file[0]);
	//ディレクトリ名取得
	String^ directory = System::IO::Path::GetDirectoryName(file[0]);

	if (extension == ".xlsx" || extension == "xls" || extension == "xlw") {
		file1->Text = file[0];
	};
	return System::Void();
}

System::Void ExcelDifCalc1::MyForm::file1_DragEnter(System::Object ^ sender, System::Windows::Forms::DragEventArgs ^ e)
{
	if (e->Data->GetDataPresent(DataFormats::FileDrop)) {
		e->Effect = DragDropEffects::All;
	}
	else {
		e->Effect = DragDropEffects::None;
	}
}

System::Void ExcelDifCalc1::MyForm::file2_DragDrop(System::Object ^ sender, System::Windows::Forms::DragEventArgs ^ e)
{
	array<String^>^ file = (array<String^>^)e->Data->GetData(DataFormats::FileDrop, false);
	//拡張子取得
	String^	extension = System::IO::Path::GetExtension(file[0]);
	//ファイル名取得
	String^ title = System::IO::Path::GetFileName(file[0]);
	//ディレクトリ名取得
	String^ directory = System::IO::Path::GetDirectoryName(file[0]);

	if (extension == ".xlsx" || extension == "xls" || extension == "xlw") {
		file2->Text = file[0];
	};
	return System::Void();
}

System::Void ExcelDifCalc1::MyForm::file2_DragEnter(System::Object ^ sender, System::Windows::Forms::DragEventArgs ^ e)
{
	if (e->Data->GetDataPresent(DataFormats::FileDrop)) {
		e->Effect = DragDropEffects::All;
	}
	else {
		e->Effect = DragDropEffects::None;
	}
}

System::Void ExcelDifCalc1::MyForm::calcButton_Click(System::Object ^ sender, System::EventArgs ^ e)
{
	String^ filePath1 = file1->Text;
	String^ filePath2 = file2->Text;

	Microsoft::Office::Interop::Excel::Application^ app_ = nullptr;

	Microsoft::Office::Interop::Excel::Workbook^ workbook = nullptr;
	Microsoft::Office::Interop::Excel::Worksheet^ worksheet = nullptr;
	Microsoft::Office::Interop::Excel::Range^ samRange = nullptr;

	Microsoft::Office::Interop::Excel::Workbook^ workbook2 = nullptr;
	Microsoft::Office::Interop::Excel::Worksheet^ worksheet2 = nullptr;
	Microsoft::Office::Interop::Excel::Range^ samRange2 = nullptr;

	Microsoft::Office::Interop::Excel::Workbook^ workbook3 = nullptr;
	Microsoft::Office::Interop::Excel::Worksheet^ worksheet3 = nullptr;
	Microsoft::Office::Interop::Excel::Range^ samRange3 = nullptr;

	Microsoft::Office::Interop::Excel::Worksheet^ copyWorksheet = nullptr;

	Microsoft::Office::Interop::Excel::Range^ allcells = nullptr;
	Microsoft::Office::Interop::Excel::Range^ allcells2 = nullptr;
	Microsoft::Office::Interop::Excel::Range^ allcells3 = nullptr;

	//String^ path1 = "C:\\Users\\chach\\Desktop\\edc1.xlsx";
	//String^ path2 = "C:\\Users\\chach\\Desktop\\edc2.xlsx";
	String^ path3 = "C:\\Users\\chach\\Desktop\\edc3.xlsx";

	//String^ path4 = "C:\\Users\\chach\\Desktop\\folder\\edc1.xlsx";

	try {
		app_ = gcnew Microsoft::Office::Interop::Excel::ApplicationClass();
		app_->Visible = false;

		//Microsoft::Office::Interop::Excel::Workbook^ calcRes = app_->Workbooks->Add(Type::Missing);

		workbook = (Microsoft::Office::Interop::Excel::Workbook^)(app_->Workbooks->Open(
			filePath1,
			//path4,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing));

		worksheet = (Microsoft::Office::Interop::Excel::Worksheet^)workbook->Worksheets[1];

		workbook3 = (Microsoft::Office::Interop::Excel::Workbook^)(app_->Workbooks->Open(
			path3,
			//calcRes->Name,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing));

		worksheet3 = (Microsoft::Office::Interop::Excel::Worksheet^)workbook3->Worksheets[1];
		/*workbook3->Worksheets->Add(
			Type::Missing,
			worksheet3,
			Type::Missing,
			Type::Missing);*/


		worksheet->Copy(Type::Missing, worksheet3);

		copyWorksheet = (Microsoft::Office::Interop::Excel::Worksheet^)workbook3->Worksheets[2];

		workbook->Close(true, file1, false);



		workbook2 = (Microsoft::Office::Interop::Excel::Workbook^)(app_->Workbooks->Open(
			//path2,
			filePath2,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing,
			Type::Missing));

		worksheet2 = (Microsoft::Office::Interop::Excel::Worksheet^)workbook2->Worksheets[1];



		int endCt = 0;
		for (int i = 1; i < copyWorksheet->Columns->Count; i++) {
			for (int j = 1; j < copyWorksheet->Rows->Count; j++) {
				//samRange = (Microsoft::Office::Interop::Excel::Range^)worksheet->Cells[i, j];
				samRange = (Microsoft::Office::Interop::Excel::Range^)copyWorksheet->Cells[i, j];
				samRange2 = (Microsoft::Office::Interop::Excel::Range^)worksheet2->Cells[i, j];
				samRange3 = (Microsoft::Office::Interop::Excel::Range^)worksheet3->Cells[i, j];

				//allcells = (Microsoft::Office::Interop::Excel::Range^) worksheet->Cells;
				allcells = (Microsoft::Office::Interop::Excel::Range^) copyWorksheet->Cells;
				allcells2 = (Microsoft::Office::Interop::Excel::Range^) worksheet2->Cells;
				allcells3 = (Microsoft::Office::Interop::Excel::Range^) worksheet3->Cells;

				if (samRange->Text == "") {
					if (endCt == 1) {
						return;
					}
					endCt++;
					break;
				}
				else {
					double res1, res2;
					double::TryParse(samRange->Text->ToString(), res1);
					double::TryParse(samRange2->Text->ToString(), res2);
					double sub = res1 - res2;
					samRange3->Value2 = sub;
					endCt = 0;
				}
			}
		}

	}
	catch (Exception^ ex) {
		MessageBox::Show(ex->ToString());
	}
	finally{
		copyWorksheet->Delete();
		workbook3->Save();
		app_->Workbooks->Close();

		System::Runtime::InteropServices::Marshal::ReleaseComObject(samRange3);
		System::Runtime::InteropServices::Marshal::ReleaseComObject(allcells3);
		System::Runtime::InteropServices::Marshal::ReleaseComObject(worksheet3);
		System::Runtime::InteropServices::Marshal::ReleaseComObject(workbook3);

		System::Runtime::InteropServices::Marshal::ReleaseComObject(samRange2);
		System::Runtime::InteropServices::Marshal::ReleaseComObject(allcells2);
		System::Runtime::InteropServices::Marshal::ReleaseComObject(worksheet2);
		System::Runtime::InteropServices::Marshal::ReleaseComObject(workbook2);

		System::Runtime::InteropServices::Marshal::ReleaseComObject(samRange);
		System::Runtime::InteropServices::Marshal::ReleaseComObject(allcells);
		System::Runtime::InteropServices::Marshal::ReleaseComObject(worksheet);
		System::Runtime::InteropServices::Marshal::ReleaseComObject(workbook);
		System::Runtime::InteropServices::Marshal::ReleaseComObject(app_);
	}

}
