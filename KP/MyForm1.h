#pragma once
using namespace System::Data::OleDb;
#include <iostream>
namespace KP {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// ������ ��� MyForm1
	/// </summary>
	public ref class MyForm1 : public System::Windows::Forms::Form
	{
	public:
		MyForm1(void)
		{
			InitializeComponent();
			//
			//TODO: �������� ��� ������������
			//
		}

		void uplBD()
		{
			dataGridView1->Rows->Clear();

			//������ �����������
			String^ connectionString = "provider=Microsoft.Jet.OleDB.4.0;Data Source=Database5.mdb";
			//���������� � ��
			OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);
			//��������� ������ � ��
			//��������� ����������
			dbConnection->Open();

			String^ query = "SELECT * FROM [warehouse]";
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
					dataGridView1->Rows->Add(dbReader[0], dbReader[1], dbReader[2], dbReader[3], dbReader[4]);
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
		~MyForm1()
		{
			if (components)
			{
				delete components;
			}
		}

	protected:


	private: System::Windows::Forms::DataGridView^ dataGridView1;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ id;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ Typeofclothing;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ color;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ quantity;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ price;
	private: System::Windows::Forms::Button^ button1;

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
			System::ComponentModel::ComponentResourceManager^ resources = (gcnew System::ComponentModel::ComponentResourceManager(MyForm1::typeid));
			this->dataGridView1 = (gcnew System::Windows::Forms::DataGridView());
			this->id = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->Typeofclothing = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->color = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->quantity = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->price = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->button1 = (gcnew System::Windows::Forms::Button());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->BeginInit();
			this->SuspendLayout();
			// 
			// dataGridView1
			// 
			this->dataGridView1->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->dataGridView1->Columns->AddRange(gcnew cli::array< System::Windows::Forms::DataGridViewColumn^  >(5) {
				this->id, this->Typeofclothing,
					this->color, this->quantity, this->price
			});
			this->dataGridView1->Location = System::Drawing::Point(12, 21);
			this->dataGridView1->Name = L"dataGridView1";
			this->dataGridView1->Size = System::Drawing::Size(630, 222);
			this->dataGridView1->TabIndex = 12;
			this->dataGridView1->CellContentClick += gcnew System::Windows::Forms::DataGridViewCellEventHandler(this, &MyForm1::dataGridView1_CellContentClick);
			// 
			// id
			// 
			this->id->AutoSizeMode = System::Windows::Forms::DataGridViewAutoSizeColumnMode::DisplayedCells;
			this->id->HeaderText = L"�";
			this->id->Name = L"id";
			this->id->Width = 43;
			// 
			// Typeofclothing
			// 
			this->Typeofclothing->HeaderText = L"�������� ����������";
			this->Typeofclothing->Name = L"Typeofclothing";
			this->Typeofclothing->Width = 300;
			// 
			// color
			// 
			this->color->HeaderText = L"����";
			this->color->Name = L"color";
			this->color->Width = 75;
			// 
			// quantity
			// 
			this->quantity->HeaderText = L"����������";
			this->quantity->Name = L"quantity";
			this->quantity->Width = 50;
			// 
			// price
			// 
			this->price->HeaderText = L"����";
			this->price->Name = L"price";
			this->price->Width = 50;
			// 
			// button1
			// 
			this->button1->Location = System::Drawing::Point(386, 250);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(256, 23);
			this->button1->TabIndex = 13;
			this->button1->Text = L"�������� ���������� ������������";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &MyForm1::button1_Click_1);
			// 
			// MyForm1
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(192)), static_cast<System::Int32>(static_cast<System::Byte>(192)),
				static_cast<System::Int32>(static_cast<System::Byte>(255)));
			this->ClientSize = System::Drawing::Size(667, 277);
			this->Controls->Add(this->button1);
			this->Controls->Add(this->dataGridView1);
			this->Icon = (cli::safe_cast<System::Drawing::Icon^>(resources->GetObject(L"$this.Icon")));
			this->Name = L"MyForm1";
			this->Text = L"���������� �� ������";
			this->Load += gcnew System::EventHandler(this, &MyForm1::MyForm1_Load);
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->EndInit();
			this->ResumeLayout(false);

		}
#pragma endregion
	private: System::Void label2_Click(System::Object^ sender, System::EventArgs^ e) {
	}
	private: System::Void MyForm1_Load(System::Object^ sender, System::EventArgs^ e) {
	}
	private: System::Void button1_Click(System::Object^ sender, System::EventArgs^ e) {
		//������ �����������
		String^ connectionString = "provider=Microsoft.Jet.OleDB.4.0;Data Source=Database5.mdb";
		//���������� � ��
		OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);
		//��������� ������ � ��
		//��������� ����������
		dbConnection->Open();

		String^ query = "SELECT * FROM [warehouse]";
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
			while (dbReader->Read()) {
				dataGridView1->Rows->Add(dbReader["id"], dbReader["Typeofclothing"], dbReader["color"], dbReader["producer"], dbReader["quantity"], dbReader["price"]);
			}
		}
		//��������� ����������
		dbReader->Close();
		dbConnection->Close();


	}
	private: System::Void dataGridView1_CellContentClick(System::Object^ sender, System::Windows::Forms::DataGridViewCellEventArgs^ e) {
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
			dataGridView1->Rows[index]->Cells[3]->Value == nullptr ||
			dataGridView1->Rows[index]->Cells[4]->Value == nullptr
			)
		{
			MessageBox::Show("�� ��� ������ �������!", "��������!");
			return;
		}

		//��������� ������
		String^ EquipmentName = dataGridView1->Rows[index]->Cells[1]->Value->ToString();
		String^ color = dataGridView1->Rows[index]->Cells[2]->Value->ToString();
		String^ quantity = dataGridView1->Rows[index]->Cells[3]->Value->ToString();
		String^ price = dataGridView1->Rows[index]->Cells[4]->Value->ToString();

		String^ connectionString = "provider=Microsoft.Jet.OleDB.4.0;Data Source=Database5.mdb";
		//���������� � ��
		OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);

		//��������� ������ � ��
		//��������� ����������
		dbConnection->Open();

		String^ query = "INSERT INTO warehouse(EquipmentName, color, quantity,price)VALUES('" + EquipmentName + "','" + color + "','" + quantity + "','" + price + "')";
		OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);//�������

		//��������� ������
		if (dbComand->ExecuteNonQuery() != 1)
			MessageBox::Show("������ ���������� �������!", "������!");
		else
			MessageBox::Show("������ ���������!", "������!");

		//��������� ����������
		dbConnection->Close();
		uplBD();
	}
	};
}
