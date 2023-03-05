#pragma once

namespace ExcelDifCalc1 {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// MyForm の概要
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			//
			//TODO: ここにコンストラクター コードを追加します
			//
		}

	protected:
		/// <summary>
		/// 使用中のリソースをすべてクリーンアップします。
		/// </summary>
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::TextBox^  file1;
	protected:
	private: System::Windows::Forms::TextBox^  file2;
	private: System::Windows::Forms::Label^  label1;
	private: System::Windows::Forms::Label^  label2;
	private: System::Windows::Forms::Button^  calcButton;

	private:
		/// <summary>
		/// 必要なデザイナー変数です。
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// デザイナー サポートに必要なメソッドです。このメソッドの内容を
		/// コード エディターで変更しないでください。
		/// </summary>
		void InitializeComponent(void)
		{
			this->file1 = (gcnew System::Windows::Forms::TextBox());
			this->file2 = (gcnew System::Windows::Forms::TextBox());
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->label2 = (gcnew System::Windows::Forms::Label());
			this->calcButton = (gcnew System::Windows::Forms::Button());
			this->SuspendLayout();
			// 
			// file1
			// 
			this->file1->AllowDrop = true;
			this->file1->Location = System::Drawing::Point(23, 33);
			this->file1->Name = L"file1";
			this->file1->Size = System::Drawing::Size(191, 19);
			this->file1->TabIndex = 0;
			this->file1->DragDrop += gcnew System::Windows::Forms::DragEventHandler(this, &MyForm::file1_DragDrop);
			this->file1->DragEnter += gcnew System::Windows::Forms::DragEventHandler(this, &MyForm::file1_DragEnter);
			// 
			// file2
			// 
			this->file2->AllowDrop = true;
			this->file2->Location = System::Drawing::Point(23, 81);
			this->file2->Name = L"file2";
			this->file2->Size = System::Drawing::Size(191, 19);
			this->file2->TabIndex = 1;
			this->file2->DragDrop += gcnew System::Windows::Forms::DragEventHandler(this, &MyForm::file2_DragDrop);
			this->file2->DragEnter += gcnew System::Windows::Forms::DragEventHandler(this, &MyForm::file2_DragEnter);
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->Location = System::Drawing::Point(23, 15);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(27, 12);
			this->label1->TabIndex = 2;
			this->label1->Text = L"file1";
			// 
			// label2
			// 
			this->label2->AutoSize = true;
			this->label2->Location = System::Drawing::Point(23, 59);
			this->label2->Name = L"label2";
			this->label2->Size = System::Drawing::Size(27, 12);
			this->label2->TabIndex = 3;
			this->label2->Text = L"file2";
			// 
			// calcButton
			// 
			this->calcButton->Location = System::Drawing::Point(23, 135);
			this->calcButton->Name = L"calcButton";
			this->calcButton->Size = System::Drawing::Size(75, 23);
			this->calcButton->TabIndex = 4;
			this->calcButton->Text = L"計算";
			this->calcButton->UseVisualStyleBackColor = true;
			this->calcButton->Click += gcnew System::EventHandler(this, &MyForm::calcButton_Click);
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 12);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(284, 261);
			this->Controls->Add(this->calcButton);
			this->Controls->Add(this->label2);
			this->Controls->Add(this->label1);
			this->Controls->Add(this->file2);
			this->Controls->Add(this->file1);
			this->Name = L"MyForm";
			this->Text = L"MyForm";
			this->Load += gcnew System::EventHandler(this, &MyForm::MyForm_Load);
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
	private: System::Void MyForm_Load(System::Object^  sender, System::EventArgs^  e);
	private: System::Void file1_DragDrop(System::Object^  sender, System::Windows::Forms::DragEventArgs^  e);
	private: System::Void file1_DragEnter(System::Object^  sender, System::Windows::Forms::DragEventArgs^  e);
private: System::Void file2_DragDrop(System::Object^  sender, System::Windows::Forms::DragEventArgs^  e);
private: System::Void file2_DragEnter(System::Object^  sender, System::Windows::Forms::DragEventArgs^  e);
private: System::Void calcButton_Click(System::Object^  sender, System::EventArgs^  e);
};
}
