# STAD Report Automation

Automating the process of generating a daily report from STAD Transaction in SAP

## The Problem

So i have to generate a report that comes from STAD transaction on SAP and it tooked like half of my day to generate it and STAD data it's only stored for 48 hours normaly. In the end this is the solution i came up with.

## Step 1: On SAP

The name of the program that runs from STAD transaction is RSSTAT26 and you can schedule it to run in background daily. I neede to export the data from the report to a file, in this case i use a text file.

![STAD transaction program name](https://github.com/kasteion/STADTRXReportAutomation/blob/master/images/stad.jpg)

First I needed to create a printer in SAP that sends all the data from a report to a text file. This is done in SPAD transaction defining a "dummy printer" that instead of sending the stream of data to a phisical printer it sends it to a command. It uses this command to send the data `/usr/bin/cat &F > /procesos/STAD300.txt`

![Printer definition](https://github.com/kasteion/STADTRXReportAutomation/blob/master/images/printer01.jpg)

![Printer command](https://github.com/kasteion/STADTRXReportAutomation/blob/master/images/printer02.jpg)

In transaction SE38 you can define variants for the RSSTAT26 program. This variants set the data you want the program to select for the report.

![SE38](https://github.com/kasteion/STADTRXReportAutomation/blob/master/images/se38.jpg)

![Variants](https://github.com/kasteion/STADTRXReportAutomation/blob/master/images/variants.jpg)

Then comes the job definition in SM36 transaction:

![SM36](https://github.com/kasteion/STADTRXReportAutomation/blob/master/images/job-definition.jpg)

And in the steps for the jobs you can define the variants and even the printer from the step:

![Job Steps](https://github.com/kasteion/STADTRXReportAutomation/blob/master/images/job-definition-steps.jpg)

When this jobs runs i get a text file with the data from the STAD Transaction daily... with that i can work to generate the report.

## Step 2: Upload text file to database

The next step consists of uploading the text file to a Microsoft SQL Server database. The data is uploaded to a table called STAD and with the help of some views the data gets prepared for the report generation.

This Loading is done in CargarReporteSTAD project. This step gets executed by the Windows Task Manager.

## Step 3: Generate Reports

After the data is uploaded i have to generate reports and clean files for another process that uses the files but can't work with the originals.

This report generation is done in GenerarReporteSTAD project. This step also gets executed by the Windows Task Manager.

## Step 4: Email it to Manager

The final report gets sended to those interested.
