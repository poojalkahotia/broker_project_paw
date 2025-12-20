# Brokerage Management System - User Guide

This guide describes how to use the Brokerage Management System to manage parties, brokers, sales, purchases, and daily accounts.

## 1. Getting Started

### Accessing the Application
1.  Open your web browser.
2.  Navigate to the application URL (e.g., `https://your-app-url.com` or `http://127.0.0.1:8000` if running locally).
3.  You will be redirected to the **Login** page.

### Login entries
*   **Username**: Enter your assigned username.
*   **Password**: Enter your secret password.
*   Click **Sign In** to access the dashboard.

---

## 2. Dashboard Overview
After logging in, you will see the **Dashboard**. This is the central hub for accessing all features.

**Main Menu Options:**
*   **Master Data**: Manage Parties, Brokers, Firms, and Items.
*   **Sales**: Create and manage sale invoices.
*   **Purchase**: Create and manage purchase invoices.
*   **Daily Page**: Record daily credit (Jama) and debit (Naame) entries.
*   **Accounts/Reports**: View ledgers, balances, and statements.

---

## 3. Master Data Setup
Before entering transactions, ensure your master records are up to date.

### Manage Parties (Buyers/Sellers)
*   Go to **Parties**.
*   **Add New**: Click "Add Party" (or similar button). Fill in details like Name, Address, Mobile, Email, and Opening Balance.
*   **Edit/Delete**: Use the action buttons next to a party in the list to update or remove them.

### Manage Brokers
*   Go to **Brokers**.
*   **Add New**: Enter Broker Name, Mobile, Email, and Opening Balances.
*   **Actions**: You can edit details or remove a broker if they are no longer active.

### Manage Firms
*   Go to **Firms**.
*   **Add New**: Enter the Firm Name. This represents your internal companies or billing entities.

### Manage Items
*   Go to **Items**.
*   **Add New**: Define the products/commodities you trade (e.g., "Cotton", "Seeds").

---

## 4. Daily Operations

### Sales Management
*   **New Sale**: Click **Sales > New Sale**.
    *   Select **Date**, **Party**, **Broker**, and **Firm**.
    *   Add **Item Details**: Weight, Rate, Packing, etc.
    *   The system calculates the total amounts automatically.
    *   Click **Save**.
*   **Update/Delete**: Navigate to the **Sales List**, find the invoice by number or date, and click **Edit** or **Delete**.
*   **Print Invoice**: In the list view, click the **PDF** icon to generate a printable invoice.

### Purchase Management
*   **New Purchase**: Click **Purchase > New Purchase**.
    *   Similar to Sales, enter Supplier (Party), Broker, Date, and Item details.
    *   Click **Save** to record the purchase.
*   **Reports**: Use the **Purchase List** to filter and view past purchases.

### Daily Page (Jama / Naame)
Use this section for daily cash flow and adjustments not covered by Sale/Purchase invoices.

*   **Navigate**: Go to **Daily Page**.
*   **Select Date**: Choose the date you are working on.
*   **Jama (Credit) Entry**:
    *   Click **Add Jama**.
    *   Select Party/Broker, enter Amount and Remark.
    *   Save.
*   **Naame (Debit) Entry**:
    *   Click **Add Naame**.
    *   Select Party/Broker, enter Amount and Remark.
    *   Save.
*   **Daily Report**: You can generate a PDF summary of the day's entries.

---

## 5. Reports & Accounts

### Sales & Purchase Reports
*   **Sale Report**: View a summary of sales over a period. Filter by Party or Broker.
*   **Purchase Report**: View a summary of purchases.

### Account Statements
*   **All Party Balance**: View a snapshot of balances for all parties.
*   **Party Statement**: detailed ledger for a specific party showing all transactions (Sales, Purchases, Cash).
*   **Broker Statement**: View commissions and transactions related to a specific broker.

---

## 6. Support
If you encounter any issues or need a password reset, please contact the system administrator.
