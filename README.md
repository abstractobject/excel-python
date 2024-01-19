This is a project trying to automate some of the tedious, yet very important and fault intolerant tasks that are possible to automate based off the master BOM our team receives from the drafting team.

So far this program just does very simple tasks like filtering, sorting, adding.

That's about as compicated as some material orders need to be, e.g. hardware orders.

Some material orders include the old funky math that has been passed down as tribal knowledge from project planner to project planner, e.g. anglematic material orders.

Eventually this program is going to be the foundation of more than just hardware orders, but that was the lowest hanging fruit at the time.


Windows install instructions:

1) Install Python 3.12 from the Microsoft Store

2) Press the start buton and type PATH. CLick "Edit the System Environment Variables", then click "Environment Variables", then under Systme Variables doubleclick PATH. In this new window click Add, and paste:

C:\Users\*******\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\LocalCache\local-packages\Python312\Scripts

******* must be replaced with your username.


3) Open Powershell and run:
    python3 -m ensurepip --upgrade

    then run 

    pip install pandas pyarrow pyexcel pyexcel-xls pyexcel-xlsx ortools

