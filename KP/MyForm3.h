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
			dataGridView1->Rows->Clear();

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
		void AddOrder(System::String^ user, System::String^ TourName, System::String^ staff)
		{
			//Строка подключение
			String^ connectionString = "provider=Microsoft.Jet.OleDB.4.0;Data Source=Database5.mdb";
			//Соединение с БД
			OleDbConnection^ dbConnection = gcnew OleDbConnection(connectionString);
			String^ TourID;
			String^ userId;
			String^ staffId;
			//Выполняем запрос к БД
			//Открываем соединение


			dbConnection->Open();
			//Команда
			String^ query = "SELECT Tours.id FROM Tours WHERE(((Tours.TourName) = '" + TourName + "')) ";
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
					TourID = dbReader[0]->ToString();
				}
			}

			//MessageBox::Show(TourID, "Ошибка!");
			query = "SELECT Customers.id FROM Customers WHERE(((Customers.FIO) = '" + user + "'))";
			dbComand = gcnew OleDbCommand(query, dbConnection);
			//Считываем данные
			dbReader = dbComand->ExecuteReader();
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
					userId = dbReader[0]->ToString();
				}
			}

			query = "SELECT Staff.id FROM Staff WHERE(((Staff.FIO) = '" + staff + "')); ";
			dbComand = gcnew OleDbCommand(query, dbConnection);
			//Считываем данные
			dbReader = dbComand->ExecuteReader();
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
					staffId = dbReader[0]->ToString();
				}
			}

			String^ data = userId + "," + TourID + "," + staffId;
			query = "INSERT INTO TourOrders (userId, tourID, staffId )VALUES (" + data + ")";

			dbComand = gcnew OleDbCommand(query, dbConnection);

			//Выполняем запрос
			if (dbComand->ExecuteNonQuery() != 1)
				MessageBox::Show("Ошибка выполнения запроса!", "Ошибка!");
			else
				MessageBox::Show("Данные добавлены!", "Готово!");

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
			System::ComponentModel::ComponentResourceManager^ resources = (gcnew System::ComponentModel::ComponentResourceManager(MyForm3::typeid));
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
			this->listBox2->Size = System::Drawing::Size(292, 186);
			this->listBox2->TabIndex = 3;
			// 
			// listBox3
			// 
			this->listBox3->FormattingEnabled = true;
			this->listBox3->Location = System::Drawing::Point(507, 263);
			this->listBox3->Name = L"listBox3";
			this->listBox3->Size = System::Drawing::Size(191, 186);
			this->listBox3->TabIndex = 4;
			// 
			// button1
			// 
			this->button1->Location = System::Drawing::Point(704, 416);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(90, 33);
			this->button1->TabIndex = 5;
			this->button1->Text = L"Добавить";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &MyForm3::button1_Click);
			// 
			// MyForm3
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->BackColor = System::Drawing::Color::FromArgb(static_cast<System::Int32>(static_cast<System::Byte>(192)), static_cast<System::Int32>(static_cast<System::Byte>(192)),
				static_cast<System::Int32>(static_cast<System::Byte>(255)));
			this->ClientSize = System::Drawing::Size(806, 461);
			this->Controls->Add(this->button1);
			this->Controls->Add(this->listBox3);
			this->Controls->Add(this->listBox2);
			this->Controls->Add(this->listBox1);
			this->Controls->Add(this->dataGridView1);
			this->Icon = (cli::safe_cast<System::Drawing::Icon^>(resources->GetObject(L"$this.Icon")));
			this->Name = L"MyForm3";
			this->Text = L"Заказы туров";
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->EndInit();
			this->ResumeLayout(false);

		}
#pragma endregion

	private: System::Void button1_Click(System::Object^ sender, System::EventArgs^ e) {

		//Проверяем данные
		if (listBox1->SelectedItem == nullptr ||
			listBox2->SelectedItem == nullptr ||
			listBox3->SelectedItem == nullptr)
		{
			MessageBox::Show("Не все данные выбраны!", "Внимание!");
			return;
		}

		AddOrder(listBox1->SelectedItem->ToString(), listBox2->SelectedItem->ToString(), listBox3->SelectedItem->ToString());
		uplBD();


	}

	};
}
