This code goes through my email for a specific email title, grabs the most recent one and downloads that excel file, then it will copy this data into a main excel workbook. It will also adjust a couple of equations in the relevant excel workbook and should date them as they are meant to be.
A big challenge was the datetime logic. I needed it to use yesterday and the day before that but Tuesdays kept causing issues for my formula as it kept wanting to input sundays date, when there is no existing sheet.