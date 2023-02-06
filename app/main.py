import io

import xlsxwriter
from fastapi import FastAPI
from fastapi.responses import StreamingResponse

from app.models import ExcelDto

app = FastAPI()


@app.post("/excel")
async def read_item(params: ExcelDto):
    params_dict = params.dict()
    positions = params_dict['positions']

    # Create an in-memory output file for the new workbook.
    output = io.BytesIO()

    writeExcelFile(output, positions)

    # Rewind the buffer.
    output.seek(0)

    headers = {
        'Content-Disposition': 'attachment; filename="filename.xlsx"',
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'  # noqa: E501
    }

    return StreamingResponse(
        output,
        headers=headers
    )


def writeExcelFile(output, positions):
    # Even though the final file will be in memory the module uses temp
    # files during assembly for efficiency. To avoid this on servers that
    # don't allow temp files, for example the Google APP Engine, set the
    # 'in_memory' Workbook() constructor option as shown in the docs.
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet('Posições')

    # Write some test data.
    headerCellFormat = workbook.add_format({'bold': True, 'font_color': 'red'})

    worksheet.write(0, 0, 'Ativo', headerCellFormat)
    worksheet.write(0, 1, 'Preço Médio', headerCellFormat)
    worksheet.write(0, 2, 'Quantidade', headerCellFormat)

    moneyFormat = workbook.add_format({'num_format': 'R$#.##0'})

    row = 1
    for position in positions:
        worksheet.write(row, 0, position['asset'])
        worksheet.write(row, 1, position['averagePrice'], moneyFormat)
        worksheet.write(row, 2, position['quantity'])
        row += 1

    # Close the workbook before sending the data.
    workbook.close()
