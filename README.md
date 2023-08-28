# async-mm-v2

## Microsoft Monthly Patch Updates Report Creation
This package creates a spreadsheet, starting a report for the updates for a specified month with information from the Catalog, MSRC Security Center, WSUS Updates, and Office Updates.

Uses FastAPI and async/await in Python to create an API that will collect all the necessary information and compile it into a spreadsheet that will be downloaded from the browser. 

NOTE: Currently only working for year 2023.

Sources:

MSRC: https://msrc.microsoft.com/update-guide/deployments

Catalog: https://www.catalog.update.microsoft.com/Search.aspx?q=

KB894199: https://support.microsoft.com/help/894199

Office: https://learn.microsoft.com/en-us/officeupdates/office-updates-msi


### Docker Guide:

Download and install Docker.

Download the async-mm-v2 directory and `cd` into the folder.

Run these commands:

`docker build -t async-mm-fastapi-v2 .`

`docker run -d --name mm-fast -p 80:80 async-mm-fastapi-v2`

Check your Docker containers and ensure it is running properly.

Head to http://localhost:80 and check for the json response.

Head to http://localhost:80/report/{year}-{month}-{name}, replacing `{year}` with `2023`, `{month}` with `1` and `{name}` with `report` if you want to download the Report for January 2023.

Head to http://localhost:80/docs for the FastAPI auto documentation.


### Manual API Guide:

Within terminal inside the `async-mm-v2` folder, run: `uvicorn src.main:app` or `uvicorn src.main:app --reload` for hot reload.

It should default to http://localhost:8000. Navigate there to check the json response. (Only port 8000 if not using Docker)

Head to  http://localhost:8000/report/2023-1-reportname to  grab a  report in the form of a file download.

### Manual Guide:

Create a `main.py` file.

Import the required modules: `import src.mm.report as report` and `import asyncio`

Install the the packages from `requirements.txt`

As seen in the example `main.py` file, create an asynchronous main function and run it with `asyncio.run(main())`

Create the Report, and then run it:

`rep = report.MonthlyReport(name='Monthly Report', year=2023, month=1)`

`await rep.run()`

Everything else happens automatically and creates an .xlsx spreadsheet in the `mm` directory with the name "NAME_YYYY-DD" and sends the FileResponse object back to the browser using that newly created file.

### Notes:

Created with Python 3.11
