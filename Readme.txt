 Written by: Troy Williams 
 Email: fenris@hotmail.com 
 
 
 First off, the most of the code is original. The rest of the code has been 
 put to gether from various sources around the internet including PSC. So if you recognize code that you wrote 
 I will be more then happy to put your name to the code. 
  
  
  
 This program is designed to download source code from PSC.. 
          - It is capable of downloading up to six files at the same time(well actually, the number is virtually unlimited, but more then 4 or 5 will piss PSC off. I found that 2 gives good results). 
          - It supports cut and pasting of urls, as well as drag and dropping urls. 
 I wrote this program because my home computer (win XP pro) could not access the PSC site for some reason. 
 I was receiving the code of the day newsletter, in which were the links to that days uploads. 
 So I put two and two together and wrote an app to download the files. 
 Along the way I discovered a number of interesting things. The urls in the newsletters are redirect links 
 So I developed a way to follow the redirect links. The program has a number of interesting features as well. 
 The urls to download are stored in a listbox, which basically acts like a que. This allows multiple winsock controls to download different files 
 When a valid download url is found (i.e. the normal sourcecode download page) it copies the html to a directory named for the <title> of the webpage. 
 This allows for a description of the code to be saved along with the zip file. 
 Also the html and file headers are saved in that directory as well as a *.url file with the download url. 
  
  
 The program also logs urls that caused problems to a log window, which allows the log to be written to disk. 
 The Url list can be saved to a file as well. 
  
  
 To use the program.... 
 Start up the program, either from the ide or from a full compile. 
 Select options from the menu and locate the directory where you want to download the files to, also enter the number of simultaneous downloads you want to have.  
 Goto File/New Downloads 
 An URL List window will popup, simply copy and past (or drag and drop) the urls into the listbox 
 When you are ready to begin downloading, click on the download button 
