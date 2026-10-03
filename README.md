# 📚 Gurukool Library – Library Management App

A console-based **Library Management System** written in Python. All library data (books, memberships, library info) is read from an Excel file, and every activity is automatically logged into a separate Excel log file. 🚀

---

## ✨ Features

### 🛒 1. Purchase a Book
- 📖 Browse by **genre → book code → book details** (name, author, price, stock, about)
- 👤 Works for **members** (login with Member ID + password) and **non-members**
- 💸 **Membership discount** applied automatically for members
- 📉 Handles **insufficient stock** – offers to buy the available copies instead
- 💳 Payment mode selection (modes come from the Excel file)
- 🔔 Auto-logs low stock (1–5 left) and out-of-stock books for restocking
- ❌ Logs books the customer **rejected** to buy

### 🎟️ 2. Purchase a Membership
- 📋 Shows all available memberships with details:
  - ⏳ Duration, 💰 Charges, 📚 Book access, 🔢 Issue limit, 🏷️ Discount
- 🆔 Generates a unique **Member ID** (format: `M-ABC123`)
- 🔑 Mobile number is used as the **password**
- 📅 Shows **expiry date & time** after purchase
- 🚫 Prevents buying a membership twice with the same name + mobile

### 📥 3. Issue a Book (Members only)
- 🔐 Secure login with Member ID + password
- 📏 Enforces the **issue limit** of the membership
- 🔒 Checks **book access** allowed in the membership plan
- 🚷 Blocks issuing the **same book twice**
- 📉 Reduces book stock on issue

### 📖 4. Read a Book (In-library reading)
- 🕘 Available **only during library opening hours**
- 🙋 Anyone can read – just enter name + mobile
- 🔔 Books are **auto-returned when the library closes**

### 📤 5. Return a Book
- 🎟️ Separate flows for **members** and **non-members (readers)**
- 🧾 Shows currently held books and lets you pick which to return
- 📈 Increases book stock on return

### 🗂️ 6. See Issued Books
- 🔐 Login required
- 📚 Lists all books currently issued to the member

### ⭐ 7. Rate Us
- 🌟 Rating from **0 to 5**
- 💬 Optional written review
- 🔁 One rating per person (name + mobile)

---

## ⚙️ Smart Background Features

- 🧵 **Background thread** runs every second to:
  - 🗑️ Remove **expired memberships** automatically
  - 🔄 Return all **reading books** to stock at closing time
- 🕐 All dates & times use **Indian Standard Time (IST)** 🇮🇳
- 🌐 Supports **multilingual genres** (e.g., Marathi – `कादंबरी`)
- 🛡️ Friendly error handling if the Excel file is missing

---

## ✅ Input Validations

- 🧑 **Name** → exactly 3 words, alphabets only, no short forms
- 📱 **Mobile** → 10 digits, numbers only, can't start with `0`
- 🆔 **Member ID** → must match the `M-XXX###` format
- 🔢 **Quantity / choices** → numbers only, no zero or negative values
- ✔️ **Yes/No prompts** → re-asked until a valid answer is given

---

## 📁 Excel Files

### 📥 Input: `Library_App.xlsx` (must be provided)

| Sheet | Columns (in order) |
|---|---|
| `books_detailes` | Genre, Book Code, Name, Author, Price, Available Copies, About |
| `membership_detailes` | Type, Name, Duration (days), Charges, Accessible Book Codes (comma-separated), Issue Limit, Discount % |
| `library_info` | Name, Place, Since, Staff, Opening Time, Closing Time, Payment Modes (comma-separated) |

### 📤 Output: `Library_App_Log.xlsx` (auto-created)

| Sheet | Purpose |
|---|---|
| `Registered_Members` / `Registered_Members_Log` | Active members / full membership history |
| `Purchased_books` | Book sales with discount & total |
| `Customer_Wants_More_books` | Requests exceeding available stock |
| `Buy_New_books` | Books with low stock (1–5) |
| `Hurry_Up_Buy_New_books` | Out-of-stock books |
| `Rejected_books` | Books customers chose not to buy |
| `Issued_Books` / `Issued_Books_Log` | Current issues / full issue history |
| `Rejected_To_Issue_Book` | Issue requests declined by member |
| `Reading` / `Reading_Log` | Current readers / full reading history |
| `Returned_Books` | All returned books |
| `Ratings_And_Reviews` | Customer ratings & reviews |

---

## 🛠️ Tech Stack

- 🐍 **Python 3.12+** (uses nested quotes inside f-strings)
- 📊 `openpyxl` – Excel read/write
- 🕰️ `datetime`, `zoneinfo` – IST date & time
- 🧵 `threading` – background cleanup task
- 📂 `os`, `time` – file checks & delays

---

## ▶️ How to Run

1. 📦 Install the dependency:
   ```bash
   pip install openpyxl
   ```
2. 📄 Place `Library_App.xlsx` in the **same folder** as the script.
3. ▶️ Run the app:
   ```bash
   python your_file_name.py
   ```
4. 🔢 Choose an option from the menu (press **Enter** to exit).

---

## 🧭 Main Menu

```
1. Purchase a Book
2. Purchase a membership
3. Issue a Book
4. Read a Book
5. Return a Book
6. See Issued Books
7. Rate Us
```

---

## 📝 Notes

- ⚠️ Close the Excel files while the app is running, otherwise saving may fail.
- 🪟 On Windows, you may also need `pip install tzdata` for timezone support.
- 🔐 Passwords are stored as plain text (mobile number) – fine for learning, not for production.

---

## 👨‍💻 Author

**Prathamesh Pawar** 🎓

Made with ❤️ and Python 🐍
