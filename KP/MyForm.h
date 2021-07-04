#pragma once
using namespace System::Data::OleDb;

namespace KP {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// ������ ��� MyForm
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			//
			//TODO: �������� ��� ������������
			//
		}
		void uplBD()
		{
			//������ �����������
			String^ connectionString = "provider=Microsoft.Jet.OleDB.4.0;Data Source=Database5.mdb";
			//���������� � ��
			OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);
			//��������� ������ � ��
			//��������� ����������
			dbConnection->Open();

			String^ query = "SELECT * FROM [Customers]";
			//�������
			OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
			//��������� ������
			OleDbDataReader^ dbReader = dbComand->ExecuteReader();
			//��������� ������
			if (dbReader->HasRows == false)
			{
				MessageBox::Show("������ ���������� ������!", "������!");
			}
			else
			{
				//��������� �������
				while (dbReader->Read())
				{
					dataGridView1->Rows->Add(dbReader[0], dbReader[1], dbReader[2], dbReader[3]);
				}
			}
			//��������� ����������
			dbReader->Close();
			dbConnection->Close();

		}
	protected:
		/// <summary>
		/// ���������� ��� ������������ �������.
		/// </summary>
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::DataGridView^ dataGridView1;
	protected:




	private: System::Windows::Forms::Button^ button1;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ id;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ FIO;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ Telephone;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ Address;


	protected:


	protected:

	private:
		/// <summary>
		/// ������������ ���������� ������������.
		/// </summary>
		System::ComponentModel::Container^ components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// ��������� ����� ��� ��������� ������������ � �� ��������� 
		/// ���������� ����� ������ � ������� ��������� ����.
		/// </summary>
		void InitializeComponent(void)
		{
			this->dataGridView1 = (gcnew System::Windows::Forms::DataGridView());
			this->button1 = (gcnew System::Windows::Forms::Button());
			this->id = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->FIO = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->Telephone = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->Address = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->BeginInit();
			this->SuspendLayout();
			// 
			// dataGridView1
			// 
			this->dataGridView1->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->dataGridView1->Columns->AddRange(gcnew cli::array< System::Windows::Forms::DataGridViewColumn^  >(4) {
				this->id, this->FIO,
					this->Telephone, this->Address
			});
			this->dataGridView1->Location = System::Drawing::Point(12, 12);
			this->dataGridView1->Name = L"dataGridView1";
			this->dataGridView1->Size = System::Drawing::Size(651, 215);
			this->dataGridView1->TabIndex = 12;
			this->dataGridView1->CellContentClick += gcnew System::Windows::Forms::DataGridViewCellEventHandler(this, &MyForm::dataGridView1_CellContentClick_1);
			// 
			// button1
			// 
			this->button1->Location = System::Drawing::Point(371, 233);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(292, 23);
			this->button1->TabIndex = 13;
			this->button1->Text = L"�������� ���������� ������������";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &MyForm::button1_Click_1);
			// 
			// id
			// 
			this->id->HeaderText = L"�";
			this->id->MinimumWidth = 10;
			this->id->Name = L"id";
			this->id->Width = 30;
			// 
			// FIO
			// 
			this->FIO->HeaderText = L"���";
			this->FIO->MinimumWidth = 70;
			this->FIO->Name = L"FIO";
			this->FIO->Width = 250;
			// 
			// Telephone
			// 
			this->Telephone->HeaderText = L"�������";
			this->Telephone->MinimumWidth = 10;
			this->Telephone->Name = L"Telephone";
			this->Telephone->Width = 150;
			// 
			// Address
			// 
			this->Address->HeaderText = L"�����";
			this->Address->MinimumWidth = 20;
			this->Address->Name = L"Address";
			this->Address->Width = 200;
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(192)), static_cast<System::Int32>(static_cast<System::Byte>(192)),
				static_cast<System::Int32>(static_cast<System::Byte>(255)));
			this->ClientSize = System::Drawing::Size(675, 268);
			this->Controls->Add(this->button1);
			this->Controls->Add(this->dataGridView1);
			this->MaximizeBox = false;
			this->MaximumSize = System::Drawing::Size(691, 307);
			this->MinimumSize = System::Drawing::Size(691, 307);
			this->Name = L"MyForm";
			this->StartPosition = System::Windows::Forms::FormStartPosition::CenterScreen;
			this->Text = L"��������� �������������� �����";
			this->Load += gcnew System::EventHandler(this, &MyForm::MyForm_Load);
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->EndInit();
			this->ResumeLayout(false);

		}
#pragma endregion
	private: System::Void dataGridView1_CellContentClick(System::Object^ sender, System::Windows::Forms::DataGridViewCellEventArgs^ e) {
	}
	private: System::Void button1_Click(System::Object^ sender, System::EventArgs^ e) {
	}


	private: System::Void label3_Click(System::Object^ sender, System::EventArgs^ e) {
	}
	private: System::Void label6_Click(System::Object^ sender, System::EventArgs^ e) {
	}
	private: System::Void MyForm_Load(System::Object^ sender, System::EventArgs^ e) {
	}
	private: System::Void dataGridView1_CellContentClick_1(System::Object^ sender, System::Windows::Forms::DataGridViewCellEventArgs^ e) {
	}
	private: System::Void listBox1_SelectedIndexChanged(System::Object^ sender, System::EventArgs^ e) {
	}
	private: System::Void button1_Click_1(System::Object^ sender, System::EventArgs^ e) {
		//����� ������ ������ ��� ����������
		if (dataGridView1->SelectedRows->Count != 1)
		{
			MessageBox::Show("�������� ���� ������ ��� ����������!", "��������!");
			return;
		}

		//������ ������ ��������� ������
		int index = dataGridView1->SelectedRows[0]->Index;

		//��������� ������
		if (dataGridView1->Rows[index]->Cells[1]->Value == nullptr ||
			dataGridView1->Rows[index]->Cells[2]->Value == nullptr ||
			dataGridView1->Rows[index]->Cells[3]->Value == nullptr)
		{
			MessageBox::Show("�� ��� ������ �������!", "��������!");
			return;
		}

		//��������� ������
		String^ FIO = dataGridView1->Rows[index]->Cells[1]->Value->ToString();
		String^ Telephone = dataGridView1->Rows[index]->Cells[2]->Value->ToString();
		String^ Address = dataGridView1->Rows[index]->Cells[3]->Value->ToString();


		String^ connectionString = "provider=Microsoft.Jet.OleDB.4.0;Data Source=Database5.mdb";
		//���������� � ��
		OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

		//��������� ������ � ��
		//��������� ����������
		dbConnection->Open();

		String^ query = "INSERT INTO Customers(FIO, Telephone, Address)VALUES('" + FIO + "','" + Telephone + "','" + Address + "')";
		OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);//�������

		//��������� ������
		if (dbComand->ExecuteNonQuery() != 1)
			MessageBox::Show("������ ���������� �������!", "������!");
		else
			MessageBox::Show("������ ���������!", "������!");

		//��������� ����������
		dbConnection->Close();

	}
	};
}
