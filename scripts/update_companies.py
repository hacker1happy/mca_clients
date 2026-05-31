import pandas as pd
import oracledb
import traceback
from datetime import datetime

# =====================================================
# CONFIG
# =====================================================

EXCEL_FILE = r"D:\FinAssure\Code-House\mca_clients\DB_COMPANIES.xlsx"

DB_CONFIG = {
    "user": "KINHERITANCE_USER",
    "password": "OurStartup1!",
    "dsn": """
    (description=
      (retry_count=20)
      (retry_delay=3)
      (address=
        (protocol=tcps)
        (port=1522)
        (host=adb.ap-mumbai-1.oraclecloud.com)
      )
      (connect_data=
        (service_name=gbad14cef94c10b_shareholdersdetails_medium.adb.oraclecloud.com)
      )
      (security=
        (ssl_server_dn_match=yes)
      )
    )
    """
}

BATCH_SIZE = 1000

conn = None

try:

    # =====================================================
    # CONNECT
    # =====================================================

    print(f"[{datetime.now()}] Connecting...")

    conn = oracledb.connect(**DB_CONFIG)
    cursor = conn.cursor()

    print("Connected successfully")

    # =====================================================
    # READ EXCEL
    # =====================================================

    df = pd.read_excel(EXCEL_FILE)

    df.columns = (
        df.columns
        .str.strip()
        .str.upper()
    )

    print("\nExcel Columns:")
    print(df.columns.tolist())

    required_cols = [
        "COMPANY_NAME",
        "G_COMPANY_NAME",
        "CIN_NUMBER",
        "SYMBOL",
        "FOLDER_NAME"
    ]

    missing_cols = [
        col for col in required_cols
        if col not in df.columns
    ]

    if missing_cols:
        raise Exception(
            f"Missing columns: {missing_cols}"
        )

    # =====================================================
    # KEEP ONLY NON EMPTY G_COMPANY_NAME
    # =====================================================

    df = df[
        df["G_COMPANY_NAME"].notna()
    ]

    df["G_COMPANY_NAME"] = (
        df["G_COMPANY_NAME"]
        .astype(str)
        .str.strip()
    )

    df = df[
        df["G_COMPANY_NAME"] != ""
    ]

    print(f"\nRows to process: {len(df)}")

    # =====================================================
    # FETCH EXISTING COMPANIES
    # =====================================================

    cursor.execute("""
        SELECT COMPANY_NAME
        FROM COMPANIES
        WHERE COMPANY_NAME IS NOT NULL
    """)

    existing_companies = {
        str(row[0]).strip().upper()
        for row in cursor.fetchall()
        if row[0]
    }

    print(
        f"Existing companies loaded: "
        f"{len(existing_companies)}"
    )

    # =====================================================
    # GET NEXT COMP_ID
    # =====================================================

    cursor.execute("""
        SELECT NVL(MAX(COMP_ID), 0)
        FROM COMPANIES
    """)

    next_comp_id = cursor.fetchone()[0] + 1

    print(
        f"Starting COMP_ID from: "
        f"{next_comp_id}"
    )

    # =====================================================
    # PREPARE DATA
    # =====================================================

    update_rows = []
    insert_rows = []
    skipped_rows = []

    for idx, row in df.iterrows():

        try:

            company_name = row["COMPANY_NAME"]

            company_name_key = (
                str(company_name).strip().upper()
                if pd.notna(company_name)
                else ""
            )

            g_company_name = (
                str(row["G_COMPANY_NAME"]).strip()
                if pd.notna(row["G_COMPANY_NAME"])
                else None
            )

            cin_number = (
                str(row["CIN_NUMBER"]).strip()
                if pd.notna(row["CIN_NUMBER"])
                else None
            )

            symbol = (
                str(row["SYMBOL"]).strip()
                if pd.notna(row["SYMBOL"])
                else None
            )

            folder_name = (
                str(row["FOLDER_NAME"]).strip()
                if pd.notna(row["FOLDER_NAME"])
                else None
            )

            if company_name_key in existing_companies:

                update_rows.append(
                    (
                        g_company_name,
                        cin_number,
                        symbol,
                        folder_name,
                        company_name
                    )
                )

            else:

                insert_rows.append(
                    (
                        next_comp_id,
                        g_company_name,
                        cin_number,
                        symbol,
                        folder_name
                    )
                )

                next_comp_id += 1

        except Exception as row_error:

            skipped_rows.append(
                f"Row {idx + 2}: {row_error}"
            )

    print(f"\nUpdate Records : {len(update_rows)}")
    print(f"Insert Records : {len(insert_rows)}")
    print(f"Skipped Records: {len(skipped_rows)}")

    # =====================================================
    # UPDATE
    # =====================================================

    update_sql = """
        UPDATE COMPANIES
        SET
            G_COMPANY_NAME = :1,
            CIN_NUMBER     = :2,
            SYMBOL         = :3,
            FOLDER_NAME    = :4
        WHERE UPPER(TRIM(COMPANY_NAME))
              = UPPER(TRIM(:5))
    """

    for i in range(0, len(update_rows), BATCH_SIZE):

        batch = update_rows[i:i+BATCH_SIZE]

        cursor.executemany(
            update_sql,
            batch
        )

        print(
            f"Updated batch "
            f"{i+1}-{i+len(batch)}"
        )

    # =====================================================
    # INSERT
    # =====================================================

    insert_sql = """
        INSERT INTO COMPANIES
        (
            COMP_ID,
            COMPANY_NAME,
            G_COMPANY_NAME,
            CIN_NUMBER,
            SYMBOL,
            FOLDER_NAME,
            PRICE,
            PRIORITY
        )
        VALUES
        (
            :1,
            NULL,
            :2,
            :3,
            :4,
            :5,
            NULL,
            NULL
        )
    """

    for i in range(0, len(insert_rows), BATCH_SIZE):

        batch = insert_rows[i:i+BATCH_SIZE]

        cursor.executemany(
            insert_sql,
            batch
        )

        print(
            f"Inserted batch "
            f"{i+1}-{i+len(batch)}"
        )

    # =====================================================
    # COMMIT
    # =====================================================

    conn.commit()

    print("\n===================================")
    print("PROCESS COMPLETED SUCCESSFULLY")
    print("===================================")
    print(f"Updated : {len(update_rows)}")
    print(f"Inserted: {len(insert_rows)}")
    print(f"Skipped : {len(skipped_rows)}")

    if skipped_rows:
        print("\nSample skipped rows:")
        for item in skipped_rows[:10]:
            print(item)

except Exception as e:

    if conn:
        conn.rollback()

    print("\nERROR OCCURRED")
    print(f"Error Type: {type(e).__name__}")
    print(f"Error Message: {e}")

    traceback.print_exc()

    print("\nTransaction rolled back.")

finally:

    try:
        cursor.close()
    except:
        pass

    try:
        conn.close()
    except:
        pass

    print("Connection closed.")