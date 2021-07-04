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
	/// Сводка для MyForm1
	/// </summary>
	public ref class MyForm1 : public System::Windows::Forms::Form
	{
	public:
		MyForm1(void)
		{
			InitializeComponent();
			//
			//TODO: добавьте код конструктора
			//
		}
		
		void uplBD()
		{
			//Строка подключение
			String^ connectionString = "provider=Microsoft.Jet.OleDB.4.0;Data Source=Database5.mdb";
			//Соединение с БД
			OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);
			//Выполняем запрос к БД
			//Открываем соединение
			dbConnection->Open();

			String^ query = "SELECT * FROM [warehouse]";
			//Команда
			OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
			//Считываем данные
			OleDbDataReader^ dbReader = dbComand->ExecuteReader();
			//Проверяем данные
			if (dbReader->HasRows == false)
			{
				MessageBox::Show("Ошибка считывания данных!", "Ошибка!");
			}
			else
			{
				//Заполняем таблицу
				while (dbReader->Read())
				{
					dataGridView1->Rows->Add(dbReader["id"], dbReader["Typeofclothing"], dbReader["color"], dbReader["producer"], dbReader["quantity"], dbReader["price"]);
				}
			}
			//Закрываем соединение
			dbReader->Close();
			dbConnection->Close();

		}
		

	protected:
		/// <summary>
		/// Освободить все используемые ресурсы.
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













	private:
		/// <summary>
		/// Обязательная переменная конструктора.
		/// </summary>
		System::ComponentModel::Container^ components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Требуемый метод для поддержки конструктора — не изменяйте 
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		void InitializeComponent(void)
		{
			this->dataGridView1 = (gcnew System::Windows::Forms::DataGridView());
			this->id = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->Typeofclothing = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->color = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->quantity = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->price = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
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
			this->dataGridView1->Size = System::Drawing::Size(630, 233);
			this->dataGridView1->TabIndex = 12;
			// 
			// id
			// 
			this->id->AutoSizeMode = System::Windows::Forms::DataGridViewAutoSizeColumnMode::DisplayedCells;
			this->id->HeaderText = L"№";
			this->id->Name = L"id";
			this->id->Width = 43;
			// 
			// Typeofclothing
			// 
			this->Typeofclothing->HeaderText = L"Название снаряжения";
			this->Typeofclothing->Name = L"Typeofclothing";
			this->Typeofclothing->Width = 200;
			// 
			// color
			// 
			this->color->HeaderText = L"Цвет";
			this->color->Name = L"color";
			this->color->Width = 75;
			// 
			// quantity
			// 
			this->quantity->HeaderText = L"Количество";
			this->quantity->Name = L"quantity";
			this->quantity->Width = 50;
			// 
			// price
			// 
			this->price->HeaderText = L"Цена";
			this->price->Name = L"price";
			this->price->Width = 50;
			// 
			// MyForm1
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->BackColor = System::Drawing::SystemColors::ButtonFace;
			this->ClientSize = System::Drawing::Size(667, 277);
			this->Controls->Add(this->dataGridView1);
			this->Name = L"MyForm1";
			this->Text = L"Снаряжение на складе";
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
		//Строка подключение
		String^ connectionString = "provider=Microsoft.Jet.OleDB.4.0;Data Source=Database5.mdb";
		//Соединение с БД
		OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);
		//Выполняем запрос к БД
		//Открываем соединение
		dbConnection->Open();

		String^ query = "SELECT * FROM [warehouse]";
		//Команда
		OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
		//Считываем данные
		OleDbDataReader^ dbReader = dbComand->ExecuteReader();
		//Проверяем данные
		if (dbReader->HasRows == false)
		{
			MessageBox::Show("Ошибка считывания данных!", "Ошибка!");
		}
		else
		{


			//Заполняем таблицу
			while (dbReader->Read()) {
				dataGridView1->Rows->Add(dbReader["id"], dbReader["Typeofclothing"], dbReader["color"], dbReader["producer"], dbReader["quantity"], dbReader["price"]);
			}
		}
		//Закрываем соединение
		dbReader->Close();
		dbConnection->Close();


	}
	};
}
