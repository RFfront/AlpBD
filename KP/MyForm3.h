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
	/// Сводка для MyForm3
	/// </summary>
	public ref class MyForm3 : public System::Windows::Forms::Form
	{
	public:
		MyForm3(void)
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
			String^ query = "SELECT TourOrders.id, Customers.FIO, Tours.TourName, Staff.FIO, TourOrders.orderDate FROM Tours INNER JOIN(Staff INNER JOIN(Customers INNER JOIN TourOrders ON Customers.id = TourOrders.userId) ON Staff.id = TourOrders.staffId) ON Tours.id = TourOrders.tourID;";
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
					dataGridView1->Rows->Add(dbReader[0], dbReader[1], dbReader[2], dbReader[3], dbReader[4]);
				}
			}
			//Закрываем соединение
			dbReader->Close();
			dbConnection->Close();

		}
		void uplBDCustomers()
		{
			//Строка подключение
			String^ connectionString = "provider=Microsoft.Jet.OleDB.4.0;Data Source=Database5.mdb";
			//Соединение с БД
			OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);
			//Выполняем запрос к БД
			//Открываем соединение
			dbConnection->Open();
			String^ query = "SELECT FIO FROM Customers ";
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
					listBox1->Items->Add(dbReader[0]);
					//dataGridView1->Rows->Add(dbReader[0], dbReader[1], dbReader[2], dbReader[3], dbReader[4]);
				}
			}
			//Закрываем соединение
			dbReader->Close();
			dbConnection->Close();

		}
		void uplBDTourName()
		{
			//Строка подключение
			String^ connectionString = "provider=Microsoft.Jet.OleDB.4.0;Data Source=Database5.mdb";
			//Соединение с БД
			OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);
			//Выполняем запрос к БД
			//Открываем соединение
			dbConnection->Open();
			String^ query = "SELECT TourName FROM Tours";
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
					listBox2->Items->Add(dbReader[0]);
				}
			}
			//Закрываем соединение
			dbReader->Close();
			dbConnection->Close();

		}
		void uplBDStaff()
		{
			//Строка подключение
			String^ connectionString = "provider=Microsoft.Jet.OleDB.4.0;Data Source=Database5.mdb";
			//Соединение с БД
			OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);
			//Выполняем запрос к БД
			//Открываем соединение
			dbConnection->Open();
			String^ query = "SELECT FIO FROM Staff ";
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
					listBox3->Items->Add(dbReader[0]);
				}
			}
			//Закрываем соединение
			dbReader->Close();
			dbConnection->Close();

		}
		void AddOrder(int userId, int clothesid, int staffId)
		{
			//Строка подключение
			String^ connectionString = "provider=Microsoft.Jet.OleDB.4.0;Data Source=Database5.mdb";
			//Соединение с БД
			OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);
			//Выполняем запрос к БД
			//Открываем соединение
			String^ data = userId.ToString() + "," + clothesid.ToString() + "," + staffId.ToString();
			dbConnection->Open();
			String^ query = "INSERT INTO Orders (userId, clothesid, staffId )VALUES (" + data + ")";
			//Команда
			OleDbCommand^ dbComand = gcnew OleDbCommand(query, dbConnection);
			//Считываем данные
			OleDbDataReader^ dbReader = dbComand->ExecuteReader();
			//Проверяем данные

			//Закрываем соединение
			dbReader->Close();
			dbConnection->Close();

		}
	protected:
		/// <summary>
		/// Освободить все используемые ресурсы.
		/// </summary>
		~MyForm3()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::DataGridView^ dataGridView1;
	protected:





	private: System::Windows::Forms::ListBox^ listBox1;

	private: System::Windows::Forms::ListBox^ listBox2;
	private: System::Windows::Forms::ListBox^ listBox3;
	private: System::Windows::Forms::Button^ button1;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ id;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ userId;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ clothesid;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ staffId;
	private: System::Windows::Forms::DataGridViewTextBoxColumn^ orderDate;






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
			this->userId = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->clothesid = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->staffId = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->orderDate = (gcnew System::Windows::Forms::DataGridViewTextBoxColumn());
			this->listBox1 = (gcnew System::Windows::Forms::ListBox());
			this->listBox2 = (gcnew System::Windows::Forms::ListBox());
			this->listBox3 = (gcnew System::Windows::Forms::ListBox());
			this->button1 = (gcnew System::Windows::Forms::Button());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->BeginInit();
			this->SuspendLayout();
			// 
			// dataGridView1
			// 
			this->dataGridView1->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->dataGridView1->Columns->AddRange(gcnew cli::array< System::Windows::Forms::DataGridViewColumn^  >(5) {
				this->id, this->userId,
					this->clothesid, this->staffId, this->orderDate
			});
			this->dataGridView1->Location = System::Drawing::Point(12, 12);
			this->dataGridView1->Name = L"dataGridView1";
			this->dataGridView1->Size = System::Drawing::Size(782, 239);
			this->dataGridView1->TabIndex = 0;
			// 
			// id
			// 
			this->id->HeaderText = L"№";
			this->id->Name = L"id";
			this->id->Width = 30;
			// 
			// userId
			// 
			this->userId->HeaderText = L"ФИО пользователя";
			this->userId->Name = L"userId";
			this->userId->Width = 200;
			// 
			// clothesid
			// 
			this->clothesid->HeaderText = L"Название тура";
			this->clothesid->Name = L"clothesid";
			this->clothesid->Width = 200;
			// 
			// staffId
			// 
			this->staffId->HeaderText = L"ФИО сопровождающего";
			this->staffId->Name = L"staffId";
			this->staffId->Width = 200;
			// 
			// orderDate
			// 
			this->orderDate->HeaderText = L"Дата заказа";
			this->orderDate->Name = L"orderDate";
			// 
			// listBox1
			// 
			this->listBox1->FormattingEnabled = true;
			this->listBox1->Location = System::Drawing::Point(12, 263);
			this->listBox1->Name = L"listBox1";
			this->listBox1->Size = System::Drawing::Size(191, 186);
			this->listBox1->TabIndex = 1;
			// 
			// listBox2
			// 
			this->listBox2->FormattingEnabled = true;
			this->listBox2->Location = System::Drawing::Point(209, 263);
			this->listBox2->Name = L"listBox2";
			this->listBox2->Size = System::Drawing::Size(191, 186);
			this->listBox2->TabIndex = 3;
			// 
			// listBox3
			// 
			this->listBox3->FormattingEnabled = true;
			this->listBox3->Location = System::Drawing::Point(406, 263);
			this->listBox3->Name = L"listBox3";
			this->listBox3->Size = System::Drawing::Size(191, 186);
			this->listBox3->TabIndex = 4;
			// 
			// button1
			// 
			this->button1->Location = System::Drawing::Point(603, 416);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(90, 33);
			this->button1->TabIndex = 5;
			this->button1->Text = L"button1";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &MyForm3::button1_Click);
			// 
			// MyForm3
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(806, 461);
			this->Controls->Add(this->button1);
			this->Controls->Add(this->listBox3);
			this->Controls->Add(this->listBox2);
			this->Controls->Add(this->listBox1);
			this->Controls->Add(this->dataGridView1);
			this->Name = L"MyForm3";
			this->Text = L"Заказы";
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->EndInit();
			this->ResumeLayout(false);

		}
#pragma endregion

	private: System::Void button1_Click(System::Object^ sender, System::EventArgs^ e) {
		//listBox1->SelectedValue
		if (listBox1->SelectedIndex != -1 || listBox2->SelectedIndex != -1 || listBox3->SelectedIndex != -1) {
			AddOrder(listBox1->SelectedIndex + 1, listBox2->SelectedIndex + 1, listBox3->SelectedIndex + 1);
			uplBD();
		}
	}

	};
}
