# Out-Of-Office
A google app script that links a google form and google spreadsheet to function as an out of office request and approval system

## How to Install:
### 1. Create google form for requests
Create a google form with questions you would like to know when someone submits a leave request. The following questions used to work with this script are as followed and in this order:

<img width="400" alt="Screen Shot 2021-09-25 at 4 09 32 PM" src="https://user-images.githubusercontent.com/31105362/134788088-de2bc948-28ec-4209-906d-3228fc8aca2a.png">
<img width="400" alt="Screen Shot 2021-09-25 at 4 09 39 PM" src="https://user-images.githubusercontent.com/31105362/134788092-abc092b2-d903-4e2a-af9f-9f58aaa4b9d5.png">

When asking for their leave dates, make sure to use the date question feature in forms. If not, you have to convert it to a date time format in the script itself.

### 2. Link and format a spreadsheet you want to store pending and archived requests
You can use the automatic generated one with the forms or link a google spreadsheet you personally created. After linking the spreadsheet sheet to the form, create a new tab/sheet on the bottom called "Archive" where you will store processed requests. Optional: Rename the linked tab with the purple icon as "Pending".

<img width="1412" alt="Screen Shot 2021-09-25 at 4 13 37 PM" src="https://user-images.githubusercontent.com/31105362/134788156-6fc2bf47-0631-4973-ab45-5ffb6ca3c07a.png">

### 3. Install the script
a) Click the app script editor and a new script file should be automatically generated for you

<img width="300" alt="Screen Shot 2021-09-25 at 4 17 02 PM" src="https://user-images.githubusercontent.com/31105362/134788221-899ae970-9eca-4768-bbd8-cef28a8f3e4a.png"><img width="600" alt="Screen Shot 2021-09-25 at 4 18 51 PM" src="https://user-images.githubusercontent.com/31105362/134788248-58479618-65cb-4ee6-8cbe-8b836bed0518.png">

b) Rename the file to "out-office-script" and delete the generated contents. Paste in all the contents from script.gs
c) 
