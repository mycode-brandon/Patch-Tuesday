import src.mm.report as report
from fastapi import FastAPI
import os
from fastapi.responses import FileResponse

app = FastAPI()


@app.get("/")
async def read_root():
    return {"Instructions": "go to http://localhost:8000/docs for more information"}


@app.get("/report/{year}-{month}-{name}")
async def main(year: int, month: int, name: str):
    rep = report.MonthlyReport(name=name, year=year, month=month)
    file_name = f"{name}_{year}-{month:02d}.xlsx"

    file_path = os.path.join(os.getcwd(), file_name)

    await rep.run()

    '''print(rep.get_deployment_api_url(skip=0))
    print(rep.get_vulnerability_api_url(skip=0))
    print(rep.get_affectedProduct_api_url(skip=0))
    print(report.get_misc_url())
    print(rep.get_office_url())

    for kb in rep.kbs:
        print(kb)
    print(len(rep.kbs))'''
    # return {"year": year, "month": month, "name": name}
    return FileResponse(path=file_path, media_type='application/octet-stream', filename=file_name)
