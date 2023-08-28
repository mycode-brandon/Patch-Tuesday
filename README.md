# Patch-Tuesday

This repository is a collection of methods to obtain and analyze the monthly patches that Microsoft releases every patch tuesday.
Contains:
- FastAPI app that can run in the browser and download an .xlsx file containing the report
- Method of creating a .xlsx file directly on the computer
- Scripts that can be imported into PowerBI to run the same report, alongside mcode files for the data sources.

## Method:

The way these scripts work is to pull data from the open API's Microsoft provides: 
- https://api.msrc.microsoft.com/sug/v2.0/en-US/deployment
- https://api.msrc.microsoft.com/sug/v2.0/en-US/vulnerability
- https://api.msrc.microsoft.com/sug/v2.0/en-US/affectedProduct

In addition, web scraping is used to pull other updates not available via an API like the Office Updates. The module BeautifulSoup is used for parsing the HTML.

Titles of each KB article are also pulled from each KB article, requiring many HTTP requests. Async support was added to aide in the speed of running this report, so all HTTP requests happen asynchronously. Previously, scraping data like the summary, description, known issues, etc, from each KB article was being gathered. However, once I got that working Microsoft completely changed their website template being used, making all the parsing useless. Instead of recreating, I decided to eliminate that portion and only pull the bare minumum information needed via scraping.

The API app is organized using an OOP methodology, and a report object instance can be created and run using the self.run method. That's all that is needed to make the app run with your specified parameters. Modules that proved extremely helpful were pydantic in creating base models for my objects, which can easily and automatically unpack json into nested objects, as well as aiohttp for the async support, and FastAPI which is the base of the api portion of the app.

The project uses type hints made available in recent releases of Python.

Later on, I found that it was a bit easier to maintain using smaller scripts meant just for making async HTTP requests, putting the json data into dataframes, and then using PowerBI for the rest. I've included the mcode for PowerBI which is how I am now cleaning the data run in these reports. After cleaning and combining the data in PowerBI, I just copy the table into Excel and format as necessary. PowerBI doesn't have the in-built capabilties to easily make async http requests, so Python is still a critical part of the process.

# APP: async-mm-v2 Guide:

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
