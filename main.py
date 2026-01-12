import pandas as pd
from decouple import config
from datetime import datetime
from sqlalchemy import create_engine, text
from sqlalchemy.exc import SQLAlchemyError


# -----------------------------
# Database connection settings
# -----------------------------
DB_USER = config('DB_USER')
DB_PASSWORD = config('DB_PASSWORD')
DB_HOST = config('DB_HOST')
DB_PORT = config('DB_PORT')
DB_NAME = config('DB_NAME')

# Create SQLAlchemy engine for accesing the database (PostgreSQL + psycopg2 driver)
DATABASE_URL = f"postgresql+psycopg2://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{DB_NAME}"
engine = create_engine(DATABASE_URL, echo=False, future=True)

try:
    # -----------------------------
    # 1. Test connection
    # -----------------------------
    with engine.connect() as conn:
        conn.execute(text("SELECT 1"))
        print("✅ Connection to PostgreSQL is successful.")

    # -----------------------------
    # 2. Read data into Pandas
    # -----------------------------
    query = '''SELECT store_id, 
                    first_name,
                    last_name,
                    email,
                    activebool,
                    create_date,
                    last_update                
    FROM customer 
    WHERE activebool = true
    order by 1,2 
    limit 10'''
    df = pd.read_sql(query, engine)
    print("\n📊 Original DataFrame:")
    print(df)
 
    # -----------------------------
    # 3. Manipulate data with Pandas
    # -----------------------------
 
    #Adding a new column current_time with default value for current timestamp
    table_name = "customer"
    new_col = 'current_time'
    current_time_val = datetime.now()
    # Add column only if it doesn't exist
    if new_col not in df.columns:
        df[new_col] = current_time_val
    print("\n🔄 Modified DataFrame:")
    print(df)
    print(f"\n💾 Data written to {table_name} table.")

    # =========================
    # Export to Excel
    # =========================
    output_file = "MB_exported_data.xlsx"
 
    try:
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        # Write query results
            df.to_excel(writer, sheet_name="Data", index=False)

        # Write the query itself in another sheet
            pd.DataFrame({"SQL Query": [query]}).to_excel(writer, sheet_name="SQL_Query", index=False)

        print(f"✅ Data and query exported successfully to '{output_file}'")
    except Exception as e:
        print(f"❌ Error exporting to Excel: {e}")
    exit(1)
     
except SQLAlchemyError as e:
    print(f"❌ Database error: {e}")
except Exception as e:
    print(f"⚠️ Unexpected error: {e}")
finally:
    # Dispose of the engine to close all connections
    engine.dispose()
    print("\n🔌 Connection closed.")
