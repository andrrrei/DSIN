# Automation of work in the DSIN project

DSIN is a database of students in need. We, a team of members of the Trade union committee of the Moscow State University, are engaged in automating the process of document generation and database synchronization to simplify work on the project. 
Tasks that we solved:
- Database synchronization. We work with more than 200 students, so it is important for us to correctly maintain a database with a lot of information about them. Every mistake costs people real money, so we decided to minimize them by building an algorithm for automatic work with the database. We have 3 of them:<br />
   1) The database of answers to google-form that students fill out and apply for admission<br />
   2) A database with all students with their current data and documents<br />
   3) A similar database maintained by the MSU Trade Union Committee<br />
   
   The algorithm of interaction between databases: after checking the documents by the project members, if they are accepted, the data is received or updated from table 1 to table 2.<br />
In addition, to avoid mistakes on the part of our team or the MSU Trade Union Committee, we regularly check tables 2 and 3 and identify discrepancies.<br />
So that this work does not take dozens of hours and does not make mistakes, an ETL system has been developed that allows data manipulation quickly and accurately.
- Generation of google documents containing tables of students who are in a difficult financial situation and claim payments. Every month we need to create several documents in a strict format to transfer them to the Faculty administration for signing, and then to the Trade Union Committee of MSU. Creating such documents manually takes a lot of time and requires special care. We solved this problem by learning how to generate such files according to tables 1) and 2).<br />
Now the execution time of our work has been reduced by 80-90% for some tasks, and the number of errors has been minimized. 
- For the comfortable operation of this system, a telegram bot has been developed that implements control over these tasks.
- Also, there was a need to regularly transfer files from Google Grive (our main working space) to Yandex Drive, which was written custom script.
- We have created a table where students of our faculty can check anonymously their inclusion status in the database

Technologies used: Google Drive API, Google Docs API, Google Spreadsheets API, Yandex drive API.

Summing up, our team managed to deal with the Google API and Yandex API to build the ETL system and put its management in telegram-bot.

At the moment, some scripts are being finalized and debagged, so they will be uploaded in the near future. We are currently working on the final debugging of this system to distribute it to other faculties of Moscow State University.



Some words about files in this repository:
1. bases_sync.py - scrips which allows to compare two databases and identify discrepancies.
2. course_heads.py - script to generate sheets which should be sent to course heads to clarify students` status.
3. enabling.py, re_enabling.py, extending.py - scripts to generate docs.
4. drive2drive.py - sctipt to transfer some files from ouк Google Drive to Yandex Drive which belongs to MSU Trade Union Committee.
5. stat_base.py - script to build and update base with students` statuses.