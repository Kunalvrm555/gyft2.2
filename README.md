# gyft2.2
### **gyft 2.2** is a feature update to gyft - *Get Your Freakin Timetable* by metaKGP available at <a href = https://github.com/metakgp/gyft> metakgp/gyft </a>


## What gyft does ? 
Fetches your timetable from ERP and provides it in a .ics file (iCalendar file format)
## What gyft2.2 does ?
Fetches your timetable from ERP and provides it in a well-structured .xlsx file (Spreadsheet)

## How to use ?

- **Step 1** - Get your Time Table from ERP : (This step is same as Step 1 in **gyft by metaKGP**)
  
  **Run the gyft script using Python (locally)**
  
  To do this, clone this repository and run the following comamands    
   ```sh
  $ pip install -r requirements.txt
  ```

  ```sh
  $ python3 gyft.py
  ```
  Enter your Roll No. , password and answer to the security question when prompted
  
  Your timetable will be saved in `data.txt`. Make any changes required in `data.txt`. Then, proceed to Step 2. 


- **Step 2** - Generate time table.xlsx
  For this step, simply run the generate_excel.py using Python
  
  This script will use 'data.txt' obtained in Step 1 to generate a spreadsheet of your timetable.
  
  After successful execution, a file named **timetable.xlsx** will be saved in your directory.

