import tkinter as tk
import tkinter.ttk as ttk
import sqlite3
from tkinter.messagebox import showinfo, showerror
import openpyxl
import pandas as pd
from PIL import Image, ImageTk

# Определение функции delete_record для удаления записи из базы данных
def delete_record(table_name, record_id_column, record_id):
    connection = sqlite3.connect("ab.db")
    cursor = connection.cursor()
    try:
        cursor.execute(f"DELETE FROM {table_name} WHERE {record_id_column}=?", (record_id,))
        connection.commit()
        connection.close()
        showinfo("Успех", "Запись успешно удалена.")
    except sqlite3.Error as e:
        showerror("Ошибка", f"Ошибка при удалении записи: {e}")

def delete_record1(table_name1, record_id_column1, record_id1):
    connection1 = sqlite3.connect("ab.db")
    cursor1 = connection1.cursor()
    try:
        cursor1.execute(f"DELETE FROM {table_name1} WHERE {record_id_column1}=?", (record_id1,))
        connection1.commit()
        connection1.close()
        showinfo1("Успех", "Запись успешно удалена.")
    except sqlite3.Error as e:
        showerror1("Ошибка", f"Ошибка при удалении записи: {e}")

def delete_record2(table_name2, record_id_column2, record_id2):
    connection2 = sqlite3.connect("ab.db")
    cursor2 = connection2.cursor()
    try:
        cursor2.execute(f"DELETE FROM {table_name2} WHERE {record_id_column2}=?", (record_id2,))
        connection2.commit()
        connection2.close()
        showinfo2("Успех", "Запись успешно удалена.")
    except sqlite3.Error as e:
        showerror2("Ошибка", f"Ошибка при удалении записи: {e}")

def delete_record3(table_name3, record_id_column3, record_id3):
    connection3 = sqlite3.connect("ab.db")
    cursor3 = connection3.cursor()
    try:
        cursor3.execute(f"DELETE FROM {table_name3} WHERE {record_id_column3}=?", (record_id3,))
        connection3.commit()
        connection3.close()
        showinfo3("Успех", "Запись успешно удалена.")
    except sqlite3.Error as e:
        showerror3("Ошибка", f"Ошибка при удалении записи: {e}")

def delete_record4(table_name4, record_id_column4, record_id4):
    connection4 = sqlite3.connect("ab.db")
    cursor4= connection4.cursor()
    try:
        cursor4.execute(f"DELETE FROM {table_name4} WHERE {record_id_column4}=?", (record_id4,))
        connection4.commit()
        connection4.close()
        showinfo4("Успех", "Запись успешно удалена.")
    except sqlite3.Error as e:
        showerror4("Ошибка", f"Ошибка при удалении записи: {e}")

def delete_record5(table_name5, record_id_column5, record_id5):
    connection5 = sqlite3.connect("ab.db")
    cursor5 = connection5.cursor()
    try:
        cursor5.execute(f"DELETE FROM {table_name5} WHERE {record_id_column5}=?", (record_id5,))
        connection5.commit()
        connection5.close()
        showinfo5("Успех", "Запись успешно удалена.")
    except sqlite3.Error as e:
        showerror5("Ошибка", f"Ошибка при удалении записи: {e}")   

def edit_record(table_name, record_id_column, record_id, new_data):
    connection = sqlite3.connect("ab.db")
    cursor = connection.cursor()
    try:
        update_query = f"UPDATE {table_name} SET "
        for col, value in new_data.items():
            update_query += f"{col} = '{value}', "
        update_query = update_query[:-2]  # Убираем лишнюю запятую и пробел в конце
        update_query += f" WHERE {record_id_column} = ?"
        cursor.execute(update_query, (record_id,))
        connection.commit()
        connection.close()
        showinfo("Успех", "Запись успешно отредактирована.")
    except sqlite3.Error as e:
        showerror("Ошибка", f"Ошибка при редактировании записи: {e}")

def edit_record1(table_name1, record_id_column1, record_id1, new_data1):
    connection1 = sqlite3.connect("ab.db")
    cursor1 = connection1.cursor()
    try:
        update_query1 = f"UPDATE {table_name1} SET "
        for col, value in new_data1.items():
            update_query1 += f"{col} = '{value}', "
        update_query1 = update_query1[:-2]  # Убираем лишнюю запятую и пробел в конце
        update_query1 += f" WHERE {record_id_column1} = ?"
        cursor1.execute(update_query1, (record_id1,))
        connection1.commit()
        connection1.close()
        showinfo("Успех", "Запись успешно отредактирована.")
    except sqlite3.Error as e:
        showerror("Ошибка", f"Ошибка при редактировании записи: {e}")   

def edit_record2(table_name2, record_id_column2, record_id2, new_data2):
    connection2 = sqlite3.connect("ab.db")
    cursor2 = connection2.cursor()
    try:
        update_query2 = f"UPDATE {table_name2} SET "
        for col, value in new_data2.items():
            update_query2 += f"{col} = '{value}', "
        update_query2 = update_query2[:-2]  # Убираем лишнюю запятую и пробел в конце
        update_query2 += f" WHERE {record_id_column2} = ?"
        cursor2.execute(update_query2, (record_id2,))
        connection2.commit()
        connection2.close()
        showinfo("Успех", "Запись успешно отредактирована.")
    except sqlite3.Error as e:
        showerror("Ошибка", f"Ошибка при редактировании записи: {e}")

def edit_record3(table_name3, record_id_column3, record_id3, new_data3):
    connection3 = sqlite3.connect("ab.db")
    cursor3 = connection3.cursor()
    try:
        update_query3 = f"UPDATE {table_name3} SET "
        for col, value in new_data3.items():
            update_query3 += f"{col} = '{value}', "
        update_query3 = update_query3[:-2]  # Убираем лишнюю запятую и пробел в конце
        update_query3 += f" WHERE {record_id_column3} = ?"
        cursor3.execute(update_query3, (record_id3,))
        connection3.commit()
        connection3.close()
        showinfo("Успех", "Запись успешно отредактирована.")
    except sqlite3.Error as e:
        showerror("Ошибка", f"Ошибка при редактировании записи: {e}")

def edit_record4(table_name4, record_id_column4, record_id4, new_data4):
    connection4 = sqlite3.connect("ab.db")
    cursor4 = connection4.cursor()
    try:
        update_query4 = f"UPDATE {table_name4} SET "
        for col, value in new_data4.items():
            update_query4 += f"{col} = '{value}', "
        update_query4 = update_query4[:-2]  # Убираем лишнюю запятую и пробел в конце
        update_query4 += f" WHERE {record_id_column4} = ?"
        cursor4.execute(update_query4, (record_id4,))
        connection4.commit()
        connection4.close()
        showinfo("Успех", "Запись успешно отредактирована.")
    except sqlite3.Error as e:
        showerror("Ошибка", f"Ошибка при редактировании записи: {e}")

def edit_record5(table_name5, record_id_column5, record_id5, new_data5):
    connection5 = sqlite3.connect("ab.db")
    cursor5 = connection5.cursor()
    try:
        update_query5 = f"UPDATE {table_name5} SET "
        for col, value in new_data5.items():
            update_query5 += f"{col} = '{value}', "
        update_query5 = update_query5[:-2]  # Убираем лишнюю запятую и пробел в конце
        update_query5 += f" WHERE {record_id_column5} = ?"
        cursor5.execute(update_query5, (record_id5,))
        connection5.commit()
        connection5.close()
        showinfo("Успех", "Запись успешно отредактирована.")
    except sqlite3.Error as e:
        showerror("Ошибка", f"Ошибка при редактировании записи: {e}") 

def export_to_excel(table_name):
    connection = sqlite3.connect("ab.db")
    cursor = connection.cursor()
    cursor.execute(f"SELECT * FROM {table_name}")
    data = cursor.fetchall()
    columns = [desc[0] for desc in cursor.description]

    # Создаем датафрейм Pandas
    df = pd.DataFrame(data, columns=columns)

    # Создаем новый Excel-файл и сохраняем данные
    excel_file = f"{table_name}.xlsx"
    df.to_excel(excel_file, index=False)

    showinfo("Успех", f"Данные экспортированы в {excel_file}")

class YourApplication:
    def __init__(self, root):
        self.root = root
        self.root.title("Абитуриент")
        self.root.geometry("500x500")

        image = Image.open("ab.jpg")  # Замените "your_image.jpg" на путь к вашему изображению
        image.thumbnail((300, 300))  # Установите желаемые размеры

        photo = ImageTk.PhotoImage(image)

        # Создание Label для отображения изображения
        image_label = tk.Label(root, image=photo)
        image_label.image = photo

        image_label.pack()

        def show_dropdown():
            dropdown = tk.Toplevel(root)
            dropdown.title("Абитуриентам")

            tables = ["Grazhdanstvo", "Specialnost", "Izuchaemyy_yazyk", "Dopolnitelnaya_informaciya"]
            for table in tables:
                button = tk.Button(dropdown, text=table, width=20, command=lambda t=table: edit_data(t))
                button.pack()

        def edit_data(table_name):
            edit_window = tk.Toplevel(root)
            edit_window.title(f"Редактировать {table_name}")

            connection = sqlite3.connect("ab.db")  # Замените "ab.db" на имя вашей базы данных
            cursor = connection.cursor()

            def update_treeview(search_query=None):
                if self.table_pr:
                    self.table_pr.delete(*self.table_pr.get_children())
                cursor.execute(f"PRAGMA table_info({table_name})")
                columns = [column[1] for column in cursor.fetchall()]
                self.table_pr["columns"] = columns
                for col in columns:
                    self.table_pr.heading(col, text=col)
                    self.table_pr.column(col, width=100)

                where_clauses = []
                params = []
                if search_query:
                    for col in columns:
                        where_clauses.append(f"{col} LIKE ?")
                        params.append(f"%{search_query}%")
    
                where_clause = " OR ".join(where_clauses) if where_clauses else "1"
    
                query = f"SELECT * FROM {table_name} WHERE {where_clause}"
                cursor.execute(query, params)
                data = cursor.fetchall()
                for row in data:
                    self.table_pr.insert("", "end", values=row)

            self.table_pr = ttk.Treeview(edit_window)
            self.table_pr.pack()
            columns = update_treeview()

            if "#0" in self.table_pr['columns']:
                self.table_pr.column("#0", width=0)
                self.table_pr.delete("#0")

            def search():
                search_query = search_entry.get()
                update_treeview(search_query)

            search_frame = tk.Frame(edit_window)
            search_frame.pack()
            search_label = tk.Label(search_frame, text="Поиск:")
            search_label.pack(side=tk.LEFT)
            search_entry = tk.Entry(search_frame)
            search_entry.pack(side=tk.LEFT)

            search_button = tk.Button(search_frame, text="Найти", command=search)
            search_button.pack(side=tk.LEFT)

            def add_record():
                add_window = tk.Toplevel(edit_window)
                add_window.title(f"Добавить запись в {table_name}")
                entry_widgets = {}
                for col in columns:
                    entry_frame = tk.Frame(add_window)
                    entry_frame.pack()
                    entry_label = tk.Label(entry_frame, text=col, width=15)
                    entry_label.pack(side=tk.LEFT)
                    entry = tk.Entry(entry_frame, width=20)
                    entry.pack(side=tk.LEFT)
                    entry_widgets[col] = entry

                def save_record():
                    data = {col: entry.get() for col, entry in entry_widgets.items()}
                    if all(data.values()):
                        columns = ', '.join(data.keys())
                        values = ', '.join([f"'{value}'" for value in data.values()])
                        cursor.execute(f"INSERT INTO {table_name} ({columns}) VALUES ({values})")
                        connection.commit()
                        update_treeview()
                        add_window.destroy()
                    else:
                        showerror("Ошибка", "Все поля должны быть заполнены")

                save_button = tk.Button(add_window, text="Сохранить запись", width=20, command=save_record)
                save_button.pack(pady=10)

            def delete_button_clicked():
                selected_item = self.table_pr.selection()
                if not selected_item:
                    showinfo("Внимание!", "Выберите запись для удаления")
                    return

                record_id = self.table_pr.item(selected_item)['values'][0]  # Получите id из второй колонки
                table_name = "abiturient"  # Здесь укажите имя таблицы
                record_id_column = "id_abiturient"  # Здесь укажите имя столбца id
                delete_record(table_name, record_id_column, record_id)
                
                record_id1= self.table_pr.item(selected_item)['values'][0]  # Получите id из второй колонки
                table_name1 = "Grazhdanstvo"  # Здесь укажите имя таблицы
                record_id_column1 = "id_grazhdanstvo"  # Здесь укажите имя столбца id
                delete_record1(table_name1, record_id_column1, record_id1)
                
                record_id2 = self.table_pr.item(selected_item)['values'][0]  # Получите id из второй колонки
                table_name2 = "Specialnost"  # Здесь укажите имя таблицы
                record_id_column2 = "id_specialnost"  # Здесь укажите имя столбца id
                delete_record2(table_name2, record_id_column2, record_id2)
                
                record_id3 = self.table_pr.item(selected_item)['values'][0]  # Получите id из второй колонки
                table_name3 = "Izuchaemyy_yazyk"  # Здесь укажите имя таблицы
                record_id_column3 = "id_izuchaemyy_yazyk"  # Здесь укажите имя столбца id
                delete_record3(table_name3, record_id_column3, record_id3)
                
                record_id4 = self.table_pr.item(selected_item)['values'][0]  # Получите id из второй колонки
                table_name4 = "Dopolnitelnaya_informaciya"  # Здесь укажите имя таблицы
                record_id_column4 = "id_informaciya"  # Здесь укажите имя столбца id
                delete_record4(table_name4, record_id_column4, record_id4)

                record_id5 = self.table_pr.item(selected_item)['values'][0]  # Получите id из второй колонки
                table_name5 = "Uchrezhdenie_obrazovaniya"  # Здесь укажите имя таблицы
                record_id_column5 = "id_uchrezhdenie"  # Здесь укажите имя столбца id
                delete_record4(table_name5, record_id_column5, record_id5)
                
                update_treeview()

            add_button = tk.Button(edit_window, text="Добавить запись", command=add_record)
            add_button.pack()

            delete_button = tk.Button(edit_window, text="Удалить запись", command=delete_button_clicked)
            delete_button.pack()

            def edit_button_clicked():
                selected_item = self.table_pr.selection()
                if not selected_item:
                    showinfo("Внимание!", "Выберите запись для редактирования")
                    return

                # Получите данные из выделенной записи
                selected_record = self.table_pr.item(selected_item)
                record_id = selected_record['values'][0]
                record_id1 = selected_record['values'][0]
                record_id2 = selected_record['values'][0]
                record_id3 = selected_record['values'][0]
                record_id4 = selected_record['values'][0]
                record_id5 = selected_record['values'][0]
                columns = self.table_pr["columns"]
                current_data = {col: selected_record['values'][i] for i, col in enumerate(columns)}

                # Укажите имя таблицы, с которой вы хотите работать
                table_name = "abiturient"
                table_name1 = "Grazhdanstvo"
                table_name2 = "Specialnost"
                table_name3 = "Izuchaemyy_yazyk"
                table_name4 = "Dopolnitelnaya_informaciya"
                table_name5 = "Uchrezhdenie_obrazovaniya"

                # Открываем окно редактирования с текущими данными
                edit_window = tk.Toplevel(root)
                edit_window.title(f"Редактировать запись")
            
                entry_widgets = {}
                for col in columns:
                    entry_frame = tk.Frame(edit_window)
                    entry_frame.pack()
                    entry_label = tk.Label(entry_frame, text=col, width=15)
                    entry_label.pack(side=tk.LEFT)
                    entry = tk.Entry(entry_frame, width=20)
                    entry.pack(side=tk.LEFT)
                    entry.insert(0, current_data[col])  # Заполните поля текущими данными
                    entry_widgets[col] = entry

                def save_edit():
                    new_data = {col: entry.get() for col, entry in entry_widgets.items()}
                    edit_record(table_name, "id_abiturient", record_id, new_data)
                    
                    new_data1 = {col: entry.get() for col, entry in entry_widgets.items()}
                    edit_record1(table_name1, "id_grazhdanstvo", record_id1, new_data1)
                    
                    new_data2 = {col: entry.get() for col, entry in entry_widgets.items()}
                    edit_record2(table_name2, "id_specialnost", record_id2, new_data2)
                    
                    new_data3 = {col: entry.get() for col, entry in entry_widgets.items()}
                    edit_record3(table_name3, "id_Izuchaemyy_yazyk", record_id3, new_data3)
                    
                    new_data4 = {col: entry.get() for col, entry in entry_widgets.items()}
                    edit_record4(table_name4, "id_informaciya", record_id4, new_data4)
                    
                    new_data5 = {col: entry.get() for col, entry in entry_widgets.items()}
                    edit_record5(table_name5, "id_uchrezhdenie", record_id5, new_data5)
                    update_treeview()
                    edit_window.destroy()

                save_button = tk.Button(edit_window, text="Сохранить изменения", width=20, command=save_edit)
                save_button.pack(pady=10)

            edit_button = tk.Button(edit_window, text="Редактировать изменения", command=edit_button_clicked)
            edit_button.pack()

            export_button = tk.Button(edit_window, text="Экспорт в Excel", command=lambda: export_to_excel(table_name))
            export_button.pack()

        abiturient_button = tk.Button(root, text="Абитуриент", width=20, command=lambda: edit_data("abiturient"))
        abiturient_button.pack(side=tk.TOP, pady=10)

        abiturientam_button = tk.Button(root, text="Абитуриентам", width=20, command=show_dropdown)
        abiturientam_button.pack(side=tk.TOP, pady=10)

        uchrezhdenie_button = tk.Button(root, text="Учреждение образования", width=20, command=lambda: edit_data("Uchrezhdenie_obrazovaniya"))
        uchrezhdenie_button.pack(side=tk.TOP, pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    app = YourApplication(root)
    root.mainloop()