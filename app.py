from flask import Flask, render_template, request, redirect
import sqlite3
from datetime import datetime
from sklearn.linear_model import LinearRegression
import re
import os
import pandas as pd
from pypdf import PdfReader

app = Flask(__name__)

DATABASE = "database.db"

UPLOAD_FOLDER = "uploads"

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


def get_db():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn

def detect_category(description):

    text = description.lower()

    if any(word in text for word in [
        "food",
        "restaurant",
        "hotel",
        "canteen",
        "cafe",
        "tea",
        "coffee",
        "swiggy",
        "zomato",
        "pizza",
        "burger",
        "mess"
    ]):
        return "Food"


    if any(word in text for word in [
        "uber",
        "ola",
        "rapido",
        "metro",
        "bus",
        "transport",
        "petrol",
        "fuel",
        "irctc"
    ]):
        return "Travel"


    if any(word in text for word in [
        "college",
        "university",
        "school",
        "course",
        "book",
        "books",
        "education",
        "udemy",
        "coursera"
    ]):
        return "Education"


    if any(word in text for word in [
        "amazon",
        "flipkart",
        "myntra",
        "shopping",
        "store"
    ]):
        return "Shopping"


    if any(word in text for word in [
        "netflix",
        "spotify",
        "movie",
        "cinema",
        "gaming",
        "game"
    ]):
        return "Entertainment"


    if any(word in text for word in [
        "hostel",
        "pg",
        "rent",
        "room"
    ]):
        return "Hostel"


    if any(word in text for word in [
        "recharge",
        "mobile",
        "airtel",
        "jio",
        "vi",
        "vodafone"
    ]):
        return "Mobile Recharge"


    return "Other"

def transaction_exists(conn, date, amount, description):
    result = conn.execute("""
        SELECT id
        FROM transactions
        WHERE date = ?
        AND amount = ?
        AND description = ?
    """, (date, amount, description)).fetchone()

    return result is not None
def parse_statement_file(filepath):
    transactions = []

    extension = os.path.splitext(filepath)[1].lower()

    # ------------------------------------------------
    # CSV FILE
    # ------------------------------------------------
    if extension == ".csv":

        df = pd.read_csv(filepath)

        df.columns = [
            str(column).strip().lower()
            for column in df.columns
        ]

        date_column = None
        description_column = None
        amount_column = None
        debit_column = None
        credit_column = None

        for column in df.columns:

            if any(word in column for word in [
                "date", "transaction date", "txn date"
            ]):
                date_column = column

            if any(word in column for word in [
                "description", "details", "remarks",
                "narration", "merchant", "payee"
            ]):
                description_column = column

            if any(word in column for word in [
                "amount", "transaction amount",
                "value"
            ]):
                amount_column = column

            if any(word in column for word in [
                "debit", "withdrawal", "paid"
            ]):
                debit_column = column

            if any(word in column for word in [
                "credit", "deposit", "received"
            ]):
                credit_column = column

        for _, row in df.iterrows():

            try:

                if date_column is None:
                    continue

                date_value = pd.to_datetime(
                    row[date_column],
                    errors="coerce"
                )

                if pd.isna(date_value):
                    continue

                date_value = date_value.strftime("%Y-%m-%d")

                if description_column:
                    description = str(
                        row[description_column]
                    ).strip()
                else:
                    description = "Imported Transaction"

                transaction_type = "expense"

                if debit_column and pd.notna(row[debit_column]):
                    amount = float(
                        str(row[debit_column])
                        .replace(",", "")
                        .replace("₹", "")
                        .strip()
                    )

                    transaction_type = "expense"

                elif credit_column and pd.notna(row[credit_column]):

                    amount = float(
                        str(row[credit_column])
                        .replace(",", "")
                        .replace("₹", "")
                        .strip()
                    )

                    transaction_type = "income"

                elif amount_column:

                    amount = float(
                        str(row[amount_column])
                        .replace(",", "")
                        .replace("₹", "")
                        .replace("+", "")
                        .replace("-", "")
                        .strip()
                    )

                    transaction_type = "expense"

                else:
                    continue

                if amount <= 0:
                    continue

                transactions.append({
                    "date": date_value,
                    "amount": amount,
                    "description": description,
                    "category": detect_category(description),
                    "type": transaction_type
                })

            except Exception:
                continue

    # ------------------------------------------------
    # EXCEL FILE
    # ------------------------------------------------
    elif extension in [".xlsx", ".xls"]:

        df = pd.read_excel(filepath)

        df.columns = [
            str(column).strip().lower()
            for column in df.columns
        ]

        date_column = None
        description_column = None
        amount_column = None

        for column in df.columns:

            if any(word in column for word in [
                "date", "transaction date", "txn date"
            ]):
                date_column = column

            if any(word in column for word in [
                "description", "details", "remarks",
                "narration", "merchant", "payee"
            ]):
                description_column = column

            if any(word in column for word in [
                "amount", "transaction amount",
                "value", "debit", "withdrawal"
            ]):
                amount_column = column

        for _, row in df.iterrows():

            try:

                if not date_column or not amount_column:
                    continue

                date_value = pd.to_datetime(
                    row[date_column],
                    errors="coerce"
                )

                if pd.isna(date_value):
                    continue

                date_value = date_value.strftime("%Y-%m-%d")

                if description_column:
                    description = str(
                        row[description_column]
                    ).strip()
                else:
                    description = "Imported Transaction"

                amount = float(
                    str(row[amount_column])
                    .replace(",", "")
                    .replace("₹", "")
                    .replace("+", "")
                    .replace("-", "")
                    .strip()
                )

                if amount <= 0:
                    continue

                transactions.append({
                    "date": date_value,
                    "amount": amount,
                    "description": description,
                    "category": detect_category(description),
                    "type": "expense"
                })

            except Exception:
                continue

    # ------------------------------------------------
    # PDF FILE
    # ------------------------------------------------
    elif extension == ".pdf":

        reader = PdfReader(filepath)

        full_text = ""

        for page in reader.pages:

            text = page.extract_text()

            if text:
                full_text += text + "\n"

        lines = full_text.splitlines()

        for line in lines:

            line = line.strip()

            if not line:
                continue

            # Detect common date formats
            date_match = re.search(
                r"\b(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})\b",
                line
            )

            if not date_match:
                date_match = re.search(
                    r"\b(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{2,4})\b",
                    line
                )

            if not date_match:
                continue

            date_text = date_match.group(1)

            # Detect amount
            amount_matches = re.findall(
                r"(?:₹|Rs\.?|INR)?\s*"
                r"([0-9]+(?:,[0-9]{3})*(?:\.[0-9]{1,2})?)",
                line,
                re.IGNORECASE
            )

            if not amount_matches:
                continue

            try:

                amount_text = amount_matches[-1]

                amount = float(
                    amount_text
                    .replace(",", "")
                )

                if amount <= 0:
                    continue

                # Convert date
                parsed_date = None

                for date_format in [
                    "%d/%m/%Y",
                    "%d-%m-%Y",
                    "%d/%m/%y",
                    "%d-%m-%y",
                    "%d %b %Y",
                    "%d %B %Y"
                ]:

                    try:

                        parsed_date = datetime.strptime(
                            date_text,
                            date_format
                        )

                        break

                    except ValueError:
                        pass

                if parsed_date is None:
                    continue

                date_value = parsed_date.strftime(
                    "%Y-%m-%d"
                )

                # Remove date and amount from description
                description = line.replace(
                    date_text,
                    ""
                )

                description = re.sub(
                    r"(?:₹|Rs\.?|INR)?\s*"
                    r"[0-9]+(?:,[0-9]{3})*(?:\.[0-9]{1,2})?",
                    "",
                    description,
                    flags=re.IGNORECASE
                )

                description = description.strip(
                    " -|:,"
                )

                if not description:
                    description = "Imported Transaction"

                # Detect income keywords
                income_words = [
                    "received",
                    "credited",
                    "credit",
                    "refund",
                    "cashback",
                    "deposit"
                ]

                transaction_type = "expense"

                if any(
                    word in description.lower()
                    for word in income_words
                ):
                    transaction_type = "income"

                transactions.append({
                    "date": date_value,
                    "amount": amount,
                    "description": description,
                    "category": detect_category(description),
                    "type": transaction_type
                })

            except Exception:
                continue

    return transactions

def create_database():
    conn = get_db()

    conn.execute("""
        CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            type TEXT NOT NULL,
            amount REAL NOT NULL,
            category TEXT NOT NULL,
            description TEXT,
            date TEXT NOT NULL
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS budgets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            month TEXT UNIQUE NOT NULL,
            amount REAL NOT NULL
        )
    """)

    conn.commit()
    conn.close()



@app.route("/")
def home():
    return render_template("home.html")

@app.route("/dashboard")
def index():

    conn = get_db()

    transactions = conn.execute("""
        SELECT *
        FROM transactions
        ORDER BY id DESC
    """).fetchall()

    income = conn.execute("""
        SELECT COALESCE(SUM(amount), 0)
        FROM transactions
        WHERE type = 'income'
    """).fetchone()[0]

    expenses = conn.execute("""
        SELECT COALESCE(SUM(amount), 0)
        FROM transactions
        WHERE type = 'expense'
    """).fetchone()[0]

    balance = income - expenses

    conn.close()

    return render_template(
        "index.html",
        transactions=transactions,
        income=income,
        expenses=expenses,
        balance=balance
    )


@app.route("/add", methods=["GET", "POST"])
def add_transaction():

    if request.method == "POST":

        transaction_type = request.form["type"]
        amount = float(request.form["amount"])
        category = request.form["category"]
        description = request.form["description"]

        date = datetime.now().strftime("%Y-%m-%d")

        conn = get_db()

        conn.execute("""
            INSERT INTO transactions
            (type, amount, category, description, date)
            VALUES (?, ?, ?, ?, ?)
        """, (
            transaction_type,
            amount,
            category,
            description,
            date
        ))

        conn.commit()
        conn.close()

        return redirect("/")

    return render_template("add.html")


@app.route("/delete/<int:id>")
def delete_transaction(id):

    conn = get_db()

    conn.execute(
        "DELETE FROM transactions WHERE id = ?",
        (id,)
    )

    conn.commit()
    conn.close()

    return redirect("/")


@app.route("/history")
def history():

    conn = get_db()

    transactions = conn.execute("""
        SELECT *
        FROM transactions
        ORDER BY date DESC, id DESC
    """).fetchall()

    conn.close()

    return render_template(
        "history.html",
        transactions=transactions
    )

@app.route("/budget", methods=["GET", "POST"])
def budget():

    current_month = datetime.now().strftime("%Y-%m")

    conn = get_db()

    if request.method == "POST":
        budget_amount = float(request.form["budget"])

        conn.execute("""
            INSERT INTO budgets (month, amount)
            VALUES (?, ?)
            ON CONFLICT(month)
            DO UPDATE SET amount = excluded.amount
        """, (current_month, budget_amount))

        conn.commit()

    budget_result = conn.execute("""
        SELECT amount
        FROM budgets
        WHERE month = ?
    """, (current_month,)).fetchone()

    monthly_budget = budget_result["amount"] if budget_result else 0

    spent_result = conn.execute("""
        SELECT COALESCE(SUM(amount), 0)
        FROM transactions
        WHERE type = 'expense'
        AND substr(date, 1, 7) = ?
    """, (current_month,)).fetchone()

    monthly_spent = spent_result[0]

    remaining = monthly_budget - monthly_spent

    category_data = conn.execute("""
        SELECT category, SUM(amount) AS total
        FROM transactions
        WHERE type = 'expense'
        AND substr(date, 1, 7) = ?
        GROUP BY category
        ORDER BY total DESC
    """, (current_month,)).fetchall()

    conn.close()

    if monthly_budget > 0:
        percentage = (monthly_spent / monthly_budget) * 100
    else:
        percentage = 0

    return render_template(
        "budget.html",
        monthly_budget=monthly_budget,
        monthly_spent=monthly_spent,
        remaining=remaining,
        percentage=percentage,
        category_data=category_data,
        current_month=current_month
    )
@app.route("/prediction")
def prediction():

    conn = get_db()

    monthly_data = conn.execute("""
        SELECT
            substr(date, 1, 7) AS month,
            SUM(amount) AS total
        FROM transactions
        WHERE type = 'expense'
        GROUP BY substr(date, 1, 7)
        ORDER BY month
    """).fetchall()

    conn.close()

    months = []
    expenses = []

    for row in monthly_data:
        months.append(row["month"])
        expenses.append(float(row["total"]))

    prediction_value = 0

    if len(expenses) == 0:

        prediction_value = 0

    elif len(expenses) == 1:

        prediction_value = expenses[0]

    else:

        X = [[i + 1] for i in range(len(expenses))]
        y = expenses

        model = LinearRegression()
        model.fit(X, y)

        next_month_number = len(expenses) + 1

        prediction_value = model.predict(
            [[next_month_number]]
        )[0]

        if prediction_value < 0:
            prediction_value = 0

    current_date = datetime.now()

    if current_date.month == 12:
        next_month = datetime(
            current_date.year + 1,
            1,
            1
        )
    else:
        next_month = datetime(
            current_date.year,
            current_date.month + 1,
            1
        )

    next_month_name = next_month.strftime("%B %Y")

    return render_template(
        "prediction.html",
        prediction=prediction_value,
        next_month=next_month_name,
        months=months,
        expenses=expenses
    )

@app.route("/add-old", methods=["GET", "POST"])
def add_old_transaction():

    if request.method == "POST":

        transaction_type = request.form["type"]

        amount = float(
            request.form["amount"]
        )

        category = request.form["category"]

        description = request.form["description"]

        date = request.form["date"]


        conn = get_db()

        conn.execute("""
            INSERT INTO transactions
            (type, amount, category, description, date)
            VALUES (?, ?, ?, ?, ?)
        """, (
            transaction_type,
            amount,
            category,
            description,
            date
        ))

        conn.commit()
        conn.close()

        return redirect("/")


    return render_template("add_old.html")

@app.route("/import-statement", methods=["GET", "POST"])
def import_statement():

    print("IMPORT STATEMENT ROUTE OPENED")
    imported_count = 0
    duplicate_count = 0
    error = None

    if request.method == "POST":

        if "statement" not in request.files:
            error = "Please select a statement file."
            return render_template(
                "import_statement.html",
                imported_count=0,
                duplicate_count=0,
                error=error
            )

        file = request.files["statement"]

        if file.filename == "":
            error = "Please select a statement file."
            return render_template(
                "import_statement.html",
                imported_count=0,
                duplicate_count=0,
                error=error
            )

        allowed_extensions = [
            ".pdf",
            ".csv",
            ".xlsx",
            ".xls"
        ]

        extension = os.path.splitext(
            file.filename
        )[1].lower()

        if extension not in allowed_extensions:

            error = (
                "Unsupported file. "
                "Please upload PDF, CSV or Excel."
            )

            return render_template(
                "import_statement.html",
                imported_count=0,
                duplicate_count=0,
                error=error
            )

        filepath = os.path.join(
            UPLOAD_FOLDER,
            file.filename
        )

        file.save(filepath)

        try:

            transactions = parse_statement_file(
                filepath
            )

            conn = get_db()

            for transaction in transactions:

                if transaction_exists(
                    conn,
                    transaction["date"],
                    transaction["amount"],
                    transaction["description"]
                ):

                    duplicate_count += 1

                else:

                    conn.execute("""
                        INSERT INTO transactions
                        (type, amount, category, description, date)
                        VALUES (?, ?, ?, ?, ?)
                    """, (
                        transaction["type"],
                        transaction["amount"],
                        transaction["category"],
                        transaction["description"],
                        transaction["date"]
                    ))

                    imported_count += 1

            conn.commit()
            conn.close()

            os.remove(filepath)

        except Exception as e:

            error = "Could not read the statement."

            if os.path.exists(filepath):
                os.remove(filepath)

        return render_template(
            "import_statement.html",
            imported_count=imported_count,
            duplicate_count=duplicate_count,
            error=error
        )

    return render_template(
        "import_statement.html",
        imported_count=0,
        duplicate_count=0,
        error=None
    )

if __name__ == "__main__":

    create_database()

    app.run(host="0.0.0.0", port=5000)