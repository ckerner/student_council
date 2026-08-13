# student_council
Tool for managing event points for student council members



Report: Students
+-----------------------+-----------------+-------+--------+
| Name                  | Email           | Grade | Points |
+==========================================================+
| john doe              | john@cusd301.or | 9     |      5 |
|-----------------------+-----------------+-------+--------|
| frederick douglas iii | fred@no.way     | 10    |      7 |
+-----------------------+-----------------+-------+--------+
 
Report: Events
+----+------------+-------------+--------+
| ID | Date       | Description | Points |
+========================================+
|  1 | 2026-03-21 | test event  |      5 |
|----+------------+-------------+--------|
|  2 | 2026-03-21 | test 2      |      2 |
+----+------------+-------------+--------+
 
Report: Attendance
+----------+------------+-----------------------+-----------------+
| Event ID | Event      | Student               | Email           |
+=================================================================+
|        1 | test event | john doe              | john@cusd301.or |
|----------+------------+-----------------------+-----------------|
|        1 | test event | frederick douglas iii | fred@no.way     |
|----------+------------+-----------------------+-----------------|
|        2 | test 2     | frederick douglas iii | fred@no.way     |
+----------+------------+-----------------------+-----------------+
 
Report: Leaderboard
+------+-----------------------+-------+--------+
| Rank | Name                  | Grade | Points |
+===============================================+
|    1 | frederick douglas iii | 10    |      7 |
|------+-----------------------+-------+--------|
|    2 | john doe              | 9     |      5 |
+------+-----------------------+-------+--------+
 
Report: Summary
+--------------------------+-------+
| Metric                   | Value |
+==================================+
| Total Students           |     2 |
|--------------------------+-------|
| Total Events             |     2 |
|--------------------------+-------|
| Total Attendance Records |     3 |
+--------------------------+-------+
 

The report.sh utility utilizes xleak for some text based reporting.
git clone https://github.com/bgreenwell/xleak.git


