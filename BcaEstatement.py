import io
import os
import tempfile
import numpy as np
import pandas as pd
import streamlit as st

try:
    from tabula.io import read_pdf
except (ImportError, ModuleNotFoundError):
    try:
        from tabula import read_pdf
    except Exception:
        read_pdf = None

from common_ui import render_page_header, render_upload_section, render_download_section


def is_currency(value):
    if pd.isna(value) or value == '':
        return False
    try:
        float(str(value).replace(',', ''))
    except ValueError:
        return False
    else:
        return True


def clean_numeric_columns(dataframe, columns):

    for column in columns:
        dataframe[column] = dataframe[column].str.replace(',', '')
        dataframe[column] = pd.to_numeric(dataframe[column], errors='coerce')
        dataframe[column] = dataframe[column].astype('float')

    return dataframe


def union_source(dataframes):

    dfs = []
    for temp_df in dataframes:

        # Split DB into new column
        temp_df[['amount', 'type']] = temp_df[4].str.extract(r'([\d,]+(?:\.\d+)?)\s*(DB|CR)?')
        temp_df = temp_df.drop(temp_df.columns[4], axis=1)

        if(len(temp_df.columns) == 7):
            # Name column and reorder
            temp_df.columns = ['date', 'desc', 'detail', 'branch', 'balance', 'amount', 'type']
            temp_df = temp_df[['date', 'desc', 'detail', 'branch', 'amount', 'type', 'balance']]

            dfs.append(temp_df)

    if dfs:
        df = pd.concat(dfs, ignore_index=True)
    else:
        df = pd.DataFrame(columns=['date', 'desc', 'detail', 'branch', 'amount', 'type', 'balance'])
    df = df.fillna(value=np.nan)
    return df


def insert_shifted_column(dataframe):

    # Add new columns with shifted values for comparison
    dataframe['prev_date'] = dataframe['date'].shift(1)
    dataframe['prev_desc'] = dataframe['desc'].shift(1)
    dataframe['prev_detail'] = dataframe['detail'].shift(1)
    dataframe['prev_branch'] = dataframe['branch'].shift(1)
    dataframe['prev_amount'] = dataframe['amount'].shift(1)
    dataframe['prev_transaction_type'] = dataframe['type'].shift(1)
    dataframe['prev_balance'] = dataframe['balance'].shift(1)

    dataframe = dataframe.fillna(value=np.nan)

    return dataframe


def extract_transactions(dataframe):

    transactions = []
    details = []
    descs = []
    temp = {}
    last_transaction = None  # Variable to track the last added transaction

    for index, row in dataframe.iterrows():

        if row['desc'] == 'S':
            transaction = {
                "date": temp['date'],
                "desc": ' | '.join(descs).strip(' | ') if descs else '',
                "detail": ' | '.join(details).strip(' | ') if details else '',
                "branch": temp['branch'],
                "amount": temp['amount'],
                "transaction_type": temp['transaction_type'] if temp['transaction_type'] == 'DB' else 'CR',
                "balance": temp['balance']
            }
            if transaction != last_transaction:  # Avoid duplicate addition
                transactions.append(transaction)
                last_transaction = transaction
            break

        if row['desc'] == 'SALDO AWAL':
            transaction = {
                "date": row['date'],
                "desc": 'SALDO AWAL',
                "detail": '',
                "branch": '',
                "amount": '',
                "transaction_type": '',
                "balance": row['balance']
            }
            transactions.append(transaction)
            continue

        if row['desc'] == 'KETE':
            continue
        
        # New Transaction
        if not pd.isna(row['amount']) and (
            pd.isna(row['prev_amount']) or
            row['amount'] != row['prev_amount'] or
            row['desc'] != row['prev_desc'] or
            row['detail'] != row['prev_detail'] or
            row['branch'] != row['prev_branch']
        ):
            # Save previous transaction
            if temp:
                transaction = {
                    "date": temp['date'],
                    "desc": ' | '.join(descs).strip(' | ') if descs else '',
                    "detail": ' | '.join(details).strip(' | ') if details else '',
                    "branch": temp['branch'],
                    "amount": temp['amount'],
                    "transaction_type": temp['transaction_type'] if temp['transaction_type'] == 'DB' else 'CR',
                    "balance": temp['balance']
                }
                if transaction != last_transaction:  # Avoid duplicate addition
                    transactions.append(transaction)
                    last_transaction = transaction
                details = []
                descs = []
                temp = {}

            temp = {
                'date': row['date'],
                'branch': row['branch'],
                'amount': row['amount'],
                'transaction_type': row['type'],
                'balance': row['balance']
            }

        if not pd.isna(row['desc']) and row['desc'].strip():
            descs.append(row['desc'])

        if not pd.isna(row['detail']) and row['detail'].strip():
            details.append(row['detail'])

    # Save the last transaction if it hasn't been added yet
    if temp:
        transaction = {
            "date": temp['date'],
            "desc": ' | '.join(descs).strip(' | ') if descs else '',
            "detail": ' | '.join(details).strip(' | ') if details else '',
            "branch": temp['branch'],
            "amount": temp['amount'],
            "transaction_type": temp['transaction_type'] if temp['transaction_type'] == 'DB' else 'CR',
            "balance": temp['balance']
        }
        if transaction != last_transaction:  # Avoid duplicate addition
            transactions.append(transaction)

    transaction_dataframe = pd.DataFrame(transactions)

    return transaction_dataframe


def calculate_balance(dataframe, init_balance):
    dataframe['balance'] = init_balance
    # Iterate over rows
    for index, row in dataframe.iterrows():
        # If transaction type is 'DB', subtract amount from balance
        if row['transaction_type'] == 'DB':
            if index == 0:
                # For the first row, subtract amount from init_balance
                dataframe.at[index, 'balance'] -= row['amount']
            else:
                # For subsequent rows, subtract amount from the previous row's balance
                dataframe.at[index, 'balance'] = dataframe.at[index - 1, 'balance'] - row['amount']
        # If transaction type is 'CR', add amount to balance
        elif row['transaction_type'] == 'CR':
            if index == 0:
                # For the first row, add amount to init_balance
                dataframe.at[index, 'balance'] += row['amount']
            else:
                # For subsequent rows, add amount to the previous row's balance
                dataframe.at[index, 'balance'] = dataframe.at[index - 1, 'balance'] + row['amount']

    return dataframe


def save_to_excel(dataframe, output_filename, sheet_name):
    if os.path.isfile(output_filename):
        writer = pd.ExcelWriter(output_filename, engine="openpyxl", mode='a', if_sheet_exists='replace')
    else:
        writer = pd.ExcelWriter(output_filename, engine="openpyxl")

    dataframe.to_excel(writer, sheet_name=sheet_name, index=False)

    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    # Format column into currency IDR
    for cell in worksheet['E']:
        cell.number_format = '_-Rp* #,##0.00_-;[Red]-Rp* #,##0.00_-;_-Rp* "-"_-;_-@_-'
    for cell in worksheet['G']:
        cell.number_format = '_-Rp* #,##0.00_-;[Red]-Rp* #,##0.00_-;_-Rp* "-"_-;_-@_-'

    # Autofit column width
    for column_cells in worksheet.columns:
        max_length = max(len(str(cell.value)) for cell in column_cells)
        if str(column_cells[0].column) in ['5', '7']:
            worksheet.column_dimensions[column_cells[0].column_letter].width = max_length + 10
        else:
            worksheet.column_dimensions[column_cells[0].column_letter].width = max_length + 2

    writer.close()

    return


def get_year_month(sheet_name):
    parts = sheet_name.split(' ', 1)  # Split into at most two parts
    if len(parts) != 2:
        raise ValueError(f"Invalid sheet name format: {sheet_name}")
    year, month_name = parts
    month_dict = {
        'JANUARI': 1, 'FEBRUARI': 2, 'MARET': 3, 'APRIL': 4,
        'MEI': 5, 'JUNI': 6, 'JULI': 7, 'AGUSTUS': 8,
        'SEPTEMBER': 9, 'OKTOBER': 10, 'NOVEMBER': 11, 'DESEMBER': 12
    }
    if month_name not in month_dict:
        raise ValueError(f"Invalid month name in sheet name: {month_name}")
    return int(year), month_dict[month_name]


def reorder_sheets(output_filename):

    wb = load_workbook(output_filename)
    sheet_names = wb.sheetnames
    sorted_sheets = sorted(sheet_names, key=get_year_month, reverse=True)
    wb._sheets.sort(key=lambda x: sorted_sheets.index(x.title))
    wb.save(output_filename)

    return

statements_folder = "statements"

def mainBcaEstatement():
    render_page_header(
        "BCA e-Statement Converter",
        "Upload file e-statement BCA lalu export hasilnya ke Excel.",
    )

    # Upload section
    render_upload_section(
        "Upload File PDF BCA",
        "Pilih satu atau lebih file PDF e-statement BCA. Sistem akan otomatis memproses dan menggabungkan data transaksi."
    )

    # File uploader
    uploaded_files = st.file_uploader(
        "Upload PDF e-statement BCA",
        type="pdf",
        accept_multiple_files=True,
        help="Pilih file PDF statement BCA",
        label_visibility="collapsed",
    )

    if uploaded_files:
        if read_pdf is None:
            st.error("Library `tabula-py` tidak tersedia atau Java belum terinstall di server/komputer.")
            return

        progress_bar = st.progress(0)
        status_text = st.empty()
        all_transactions = []

        for i, uploaded_file in enumerate(uploaded_files):
            status_text.text(f"Memproses file {i+1} dari {len(uploaded_files)}: {uploaded_file.name}")
            progress_bar.progress((i) / len(uploaded_files))

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_path = tmp_file.name

            try:
                # Process the uploaded file
                header_data = read_pdf(tmp_path, area=(70, 315, 141, 548), pages='1', pandas_options={'header': None, 'dtype': str}, force_subprocess=True)
                if not header_data:
                    st.warning(f"Tidak dapat membaca header dari {uploaded_file.name}")
                    continue
                header_dataframe = header_data[0]
                
                # Dynamic extraction of periode/rekening could be improved but keeping existing logic for now
                try:
                    periode = header_dataframe.loc[header_dataframe[0] == 'PERIODE', 2].values[0]
                    periode = ' '.join(reversed(periode.split()))
                    no_rekening = header_dataframe.loc[header_dataframe[0] == 'NO. REKENING', 2].values[0]
                except:
                    pass

                dataframes = read_pdf(tmp_path, area=(231, 25, 797, 577), columns=[86, 184, 300, 340, 467], pages='all', pandas_options={'header': None, 'dtype': str}, force_subprocess=True)
                if not dataframes:
                    st.warning(f"Tidak dapat membaca tabel transaksi dari {uploaded_file.name}")
                    continue

                init_balance_match = dataframes[0].loc[dataframes[0][1] == 'SALDO AWAL', 5]
                if not init_balance_match.empty:
                    init_balance = float(str(init_balance_match.values[0]).replace(',', ''))
                else:
                    init_balance = 0.0

                df = union_source(dataframes)
                df = clean_numeric_columns(df, ['amount', 'balance'])
                df = insert_shifted_column(df)

                transaction_dataframe = extract_transactions(df)
                if 'balance' in transaction_dataframe.columns:
                    transaction_dataframe = transaction_dataframe.drop('balance', axis=1)
                transaction_dataframe = calculate_balance(transaction_dataframe, init_balance)

                transaction_dataframe['source_file'] = uploaded_file.name
                all_transactions.append(transaction_dataframe)
            except Exception as e:
                st.error(f"Gagal memproses file {uploaded_file.name}: {str(e)}")
            finally:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)

        progress_bar.progress(1.0)
        status_text.text("✅ Semua file berhasil diproses!")

        if all_transactions:
            # Combine all transactions
            global_dataframe = pd.concat(all_transactions, ignore_index=True)

            # Summary metrics
            total_transactions = len(global_dataframe)
            total_debit = global_dataframe[global_dataframe['transaction_type'] == 'DB']['amount'].sum()
            total_credit = global_dataframe[global_dataframe['transaction_type'] == 'CR']['amount'].sum()

            st.markdown("---")
            st.subheader("📊 Ringkasan Transaksi")

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Transaksi", f"{total_transactions:,}")
            with col2:
                st.metric("Total Debit", f"Rp {total_debit:,.0f}")
            with col3:
                st.metric("Total Kredit", f"Rp {total_credit:,.0f}")

            # Display the dataframe
            st.markdown("---")
            st.subheader("📋 Data Transaksi")
            st.dataframe(global_dataframe, use_container_width=True)

            # Download section
            render_download_section()

            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                global_dataframe.to_excel(writer, sheet_name="All Transactions", index=False)
            
            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                st.download_button(
                    label="📥 Download Excel File",
                    data=excel_buffer.getvalue(),
                    file_name="BCA_Statements_Combined.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with col_dl2:
                st.download_button(
                    label="📥 Download CSV",
                    data=global_dataframe.to_csv(index=False).encode('utf-8'),
                    file_name="BCA_Statements_Combined.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

if __name__ == "__main__":
    mainBcaEstatement()
