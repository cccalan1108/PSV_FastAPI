import os
import shutil
import io
import tempfile
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
import pandas as pd

from calc2data import convert_calc_to_data_sheet 
from data2calc import convert_data_to_calc_sheet 

app = FastAPI(
    title="PSV Excel Processing API",
    description="API for converting PSV data between Calculation Sheet and Data Sheet formats.",
    version="1.0.0",
)

# --- Configuration ---
DEFAULT_DATA_SHEET_TEMPLATE_PATH = "Data Sheet.xlsm"
DEFAULT_CALC_SHEET_TEMPLATE_PATH = "Calculation Sheet.xlsm"

TEMP_DIR = Path(tempfile.gettempdir()) / "fastapi_excel_processor_temp"
TEMP_DIR.mkdir(parents=True, exist_ok=True)

@app.post("/calc2data/", summary="Convert Calculation Sheet to Data Sheet")
async def calc2data_endpoint(calc_sheet_file: UploadFile = File(..., description="The Calculation Sheet Excel file (.xlsm)")):
    """
    Processes an uploaded Calculation Sheet Excel file and fills a Data Sheet template.
    Returns the filled Data Sheet Excel file.
    """
    temp_calc_sheet_path = None
    output_stream = io.BytesIO() 

    try:
        temp_calc_sheet_path = TEMP_DIR / f"uploaded_calc_{calc_sheet_file.filename}"
        with temp_calc_sheet_path.open("wb") as buffer:
            shutil.copyfileobj(calc_sheet_file.file, buffer)

        try:
            # Assumes 'PSV' is the sheet name in Calculation Sheet for calc2data conversion
            calc_df = pd.read_excel(temp_calc_sheet_path, sheet_name="PSV", header=None)
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Error reading Calculation Sheet: {e}. Ensure 'PSV' sheet exists and format is correct.")

        result_from_calc2data = convert_calc_to_data_sheet(calc_df, DEFAULT_DATA_SHEET_TEMPLATE_PATH, output_stream)
        
        if not result_from_calc2data:
            raise HTTPException(status_code=500, detail="Excel conversion (Calc to Data) failed. Check server logs for details.")
        
        output_stream.seek(0)

        return StreamingResponse(
            output_stream, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
            headers={"Content-Disposition": f"attachment; filename=Data_Sheet_filled_{Path(calc_sheet_file.filename).stem}.xlsm"}
        )

    except HTTPException as e:
        raise e 
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail=f"Required template file not found at '{DEFAULT_DATA_SHEET_TEMPLATE_PATH}'.")
    except Exception as e:
        print(f"An unexpected error occurred in /calc2data/: {e}") 
        raise HTTPException(status_code=500, detail=f"An internal server error occurred: {e}")
    finally:
        if temp_calc_sheet_path and temp_calc_sheet_path.exists():
            temp_calc_sheet_path.unlink()

@app.post("/data2calc/", summary="Convert Data Sheet to Calculation Sheet")
async def data2calc_endpoint(data_sheet_file: UploadFile = File(..., description="The Data Sheet Excel file (.xlsm)")):
    """
    Processes an uploaded Data Sheet Excel file and fills a Calculation Sheet template.
    Returns the filled Calculation Sheet Excel file.
    """
    temp_data_sheet_path = None
    output_stream = io.BytesIO() 

    try:
        temp_data_sheet_path = TEMP_DIR / f"uploaded_data_{data_sheet_file.filename}"
        with temp_data_sheet_path.open("wb") as buffer:
            shutil.copyfileobj(data_sheet_file.file, buffer)

        try:
            data_df = pd.read_excel(temp_data_sheet_path, sheet_name="FORM", header=None)
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Error reading Data Sheet: {e}. Ensure 'FORM' sheet exists and format is correct.")

        try:
            calc_sheet_template_df = pd.read_excel(DEFAULT_CALC_SHEET_TEMPLATE_PATH, sheet_name="PSV", header=None)
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Error reading Calculation Sheet template: {e}. Ensure '{DEFAULT_CALC_SHEET_TEMPLATE_PATH}' with sheet 'PSV' exists and is accessible.")

        result_df = convert_data_to_calc_sheet(data_df, calc_sheet_template_df)
        
        if not isinstance(result_df, pd.DataFrame):
            raise HTTPException(status_code=500, detail="Excel conversion (Data to Calc) failed: Core logic did not return a DataFrame.")

        try:
            result_df.to_excel(output_stream, index=False, header=False, engine='openpyxl')
        except Exception as e:
            raise HTTPException(status_code=500, detail=f"Error writing converted DataFrame to Excel stream: {e}")

        output_stream.seek(0)

        return StreamingResponse(
            output_stream, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
            headers={"Content-Disposition": f"attachment; filename=Calculation_Sheet_filled_{Path(data_sheet_file.filename).stem}.xlsx"}
        )

    except HTTPException as e:
        raise e
    except FileNotFoundError:
        raise HTTPException(status_code=404, detail=f"Required template file not found at '{DEFAULT_CALC_SHEET_TEMPLATE_PATH}'.")
    except Exception as e:
        print(f"An unexpected error occurred in /data2calc/: {e}")
        raise HTTPException(status_code=500, detail=f"An internal server error occurred: {e}")
    finally:
        if temp_data_sheet_path and temp_data_sheet_path.exists():
            temp_data_sheet_path.unlink()