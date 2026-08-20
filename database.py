"""
Подключение к БД и все SQL-запросы.
"""

import pyodbc

from constants import ALL_KAFEDRY, ALL_GODY


# ---------- Подключение ----------

def get_connection(server: str, database: str, auth_type: str = "windows",
                   username: str = "", password: str = ""):
    if auth_type == "windows":
        conn_str = (
            "DRIVER={ODBC Driver 17 for SQL Server};"
            f"SERVER={server};"
            f"DATABASE={database};"
            "Trusted_Connection=yes;"
        )
    else:
        conn_str = (
            "DRIVER={ODBC Driver 17 for SQL Server};"
            f"SERVER={server};"
            f"DATABASE={database};"
            f"UID={username};"
            f"PWD={password};"
        )
    try:
        return pyodbc.connect(conn_str)
    except pyodbc.Error:
        if auth_type == "windows":
            conn_str_fallback = (
                "DRIVER={SQL Server};"
                f"SERVER={server};"
                f"DATABASE={database};"
                "Trusted_Connection=yes;"
            )
        else:
            conn_str_fallback = (
                "DRIVER={SQL Server};"
                f"SERVER={server};"
                f"DATABASE={database};"
                f"UID={username};"
                f"PWD={password};"
            )
        return pyodbc.connect(conn_str_fallback)


def test_connection(server: str, database: str, auth_type: str = "windows",
                    username: str = "", password: str = "") -> tuple:
    try:
        conn = get_connection(server, database, auth_type, username, password)
        conn.close()
        return True, "Подключение успешно!"
    except pyodbc.Error as e:
        error_msg = str(e)
        if "SQL Server" in error_msg or "server" in error_msg.lower():
            return False, "Не удалось подключиться к серверу. Проверьте имя сервера."
        elif "database" in error_msg.lower():
            return False, "База данных не найдена."
        elif "login" in error_msg.lower() or "password" in error_msg.lower():
            return False, "Неверное имя пользователя или пароль."
        else:
            return False, f"Ошибка подключения:\n{error_msg}"
    except Exception as e:
        return False, f"Неизвестная ошибка:\n{str(e)}"


# ---------- Справочники ----------

def fetch_kafedry(conn):
    cursor = conn.cursor()
    cursor.execute("""
        SELECT DISTINCT к.Название FROM dbo.Кафедры к
        WHERE к.Удалена = 0 ORDER BY к.Название
    """)
    names = [row[0] for row in cursor.fetchall() if row[0]]
    cursor.close()
    return names


def fetch_prepodavateli(conn, kafedra: str):
    cursor = conn.cursor()
    if kafedra and kafedra != ALL_KAFEDRY:
        cursor.execute("""
            SELECT DISTINCT н.КодПреподавателя, н.ФИОПреподавателя
            FROM dbo.Нагрузка н
            JOIN dbo.Кафедры к ON к.Код = н.КодКафедры
            WHERE (н.ДляУдаления IS NULL OR н.ДляУдаления = 0)
              AND к.Название = ? AND н.ФИОПреподавателя IS NOT NULL
            ORDER BY н.ФИОПреподавателя
        """, [kafedra])
    else:
        cursor.execute("""
            SELECT DISTINCT н.КодПреподавателя, н.ФИОПреподавателя
            FROM dbo.Нагрузка н
            WHERE (н.ДляУдаления IS NULL OR н.ДляУдаления = 0)
              AND н.ФИОПреподавателя IS NOT NULL
            ORDER BY н.ФИОПреподавателя
        """)
    result = [(row[0], row[1]) for row in cursor.fetchall()]
    cursor.close()
    return result


def fetch_gody(conn):
    cursor = conn.cursor()
    cursor.execute("""
        SELECT DISTINCT н.УчебныйГод FROM dbo.Нагрузка н
        WHERE (н.ДляУдаления IS NULL OR н.ДляУдаления = 0)
          AND н.УчебныйГод IS NOT NULL AND н.УчебныйГод <> ''
        ORDER BY н.УчебныйГод DESC
    """)
    years = [row[0] for row in cursor.fetchall() if row[0]]
    cursor.close()
    return years


def fetch_vidy_zanyatiy(conn):
    """Список видов занятий, встречающихся в нагрузке (для фильтра)."""
    cursor = conn.cursor()
    cursor.execute("""
        SELECT DISTINCT н.ВидЗанятий FROM dbo.Нагрузка н
        WHERE (н.ДляУдаления IS NULL OR н.ДляУдаления = 0)
          AND н.ВидЗанятий IS NOT NULL AND н.ВидЗанятий <> ''
        ORDER BY н.ВидЗанятий
    """)
    values = [row[0] for row in cursor.fetchall() if row[0]]
    cursor.close()
    return values


def fetch_vidy_kontrolya(conn):
    """Список видов контроля, встречающихся в нагрузке (для фильтра)."""
    cursor = conn.cursor()
    cursor.execute("""
        SELECT DISTINCT н.ВидКонтроля FROM dbo.Нагрузка н
        WHERE (н.ДляУдаления IS NULL OR н.ДляУдаления = 0)
          AND н.ВидКонтроля IS NOT NULL AND н.ВидКонтроля <> ''
        ORDER BY н.ВидКонтроля
    """)
    values = [row[0] for row in cursor.fetchall() if row[0]]
    cursor.close()
    return values


# ---------- Данные нагрузки ----------

def fetch_nagruzka(conn, kod_prepodavatelya: int, god: str,
                    vidy_zanyatiy=None, vidy_kontrolya=None, finance=None):
    """
    vidy_zanyatiy / vidy_kontrolya: список конкретных значений для фильтра
        (IN (...)) или None/[] - фильтр не применяется.
    finance: "budget" - только строки с бюджетным финансированием (Бюджет > 0),
             "vneb"   - только строки с внебюджетным (Внебюджет > 0),
             None     - без фильтра по финансированию.
    """
    cursor = conn.cursor()
    query = """
        SELECT н.НомерСтроки, н.Группа, н.Блок, н.Дисциплина, н.ВидЗанятий,
               н.Курс, н.Семестр, н.ВидКонтроля, н.Студентов, н.НомерПотока,
               н.ИндикаторПотока, н.НагрузкаАуд, н.НагрузкаДр,
               COALESCE(н.Часов, н.НагрузкаАуд + н.НагрузкаДр) AS Итого,
               н.Бюджет, н.Внебюджет, к.Название AS Кафедра
        FROM dbo.Нагрузка н
        LEFT JOIN dbo.Кафедры к ON к.Код = н.КодКафедры
        WHERE (н.ДляУдаления IS NULL OR н.ДляУдаления = 0)
          AND н.КодПреподавателя = ?
    """
    params = [kod_prepodavatelya]

    if god and god != ALL_GODY:
        query += " AND н.УчебныйГод = ?"
        params.append(str(god))

    if vidy_zanyatiy:
        placeholders = ",".join(["?"] * len(vidy_zanyatiy))
        query += f" AND н.ВидЗанятий IN ({placeholders})"
        params.extend(vidy_zanyatiy)

    if vidy_kontrolya:
        placeholders = ",".join(["?"] * len(vidy_kontrolya))
        query += f" AND н.ВидКонтроля IN ({placeholders})"
        params.extend(vidy_kontrolya)

    if finance == "budget":
        query += " AND ISNULL(н.Бюджет, 0) > 0"
    elif finance == "vneb":
        query += " AND ISNULL(н.Внебюджет, 0) > 0"

    query += " ORDER BY н.НомерСтроки"
    cursor.execute(query, params)
    columns = [col[0] for col in cursor.description]
    rows = cursor.fetchall()
    cursor.close()
    return columns, rows


def fetch_teacher_info(conn, kod_prepodavatelya: int):
    cursor = conn.cursor()
    cursor.execute("SELECT ФИО, Должность FROM dbo.Преподаватели WHERE Код = ?", [kod_prepodavatelya])
    row = cursor.fetchone()
    fio = row[0] if row else ""
    doljnost_text = row[1] if row and row[1] else ""

    doljnost = ""
    sovmestitelstvo = ""
    cursor.execute("""
        SELECT TOP 1 сд.НазваниеДолжности, пк.Совместительство
        FROM dbo.ПреподавателиКафедры пк
        LEFT JOIN dbo.СправочникДолжности сд ON сд.Код = пк.КодДолжности
        WHERE пк.КодПреподавателя = ? AND (пк.isDelete IS NULL OR пк.isDelete = 0)
        ORDER BY пк.УчебныйГод DESC
    """, [kod_prepodavatelya])
    row2 = cursor.fetchone()
    if row2:
        if row2[0]: doljnost = row2[0]
        if row2[1]: sovmestitelstvo = row2[1]
    if not doljnost: doljnost = doljnost_text
    cursor.close()
    return {"ФИО": fio, "Должность": doljnost, "Условия": sovmestitelstvo}


def fetch_kafedra_i_fakultet_info(conn, kafedra_name: str):
    if not kafedra_name:
        return {"Заведующий": "", "Декан": ""}
    cursor = conn.cursor()
    try:
        cursor.execute("""
            SELECT ISNULL(к.ЗавКафедрой, '') AS Заведующий,
                   ISNULL(ф.Декан, '') AS Декан
            FROM dbo.Кафедры к
            LEFT JOIN dbo.Факультеты ф ON ф.Код = к.Код_Факультета
            WHERE к.Название = ? AND (к.Удалена IS NULL OR к.Удалена = 0)
        """, [kafedra_name])
        row = cursor.fetchone()
        if row:
            return {"Заведующий": (row[0] or "").strip(), "Декан": (row[1] or "").strip()}
    except pyodbc.Error as e:
        print(f"Ошибка получения заведующего/декана: {e}")
    finally:
        cursor.close()
    return {"Заведующий": "", "Декан": ""}