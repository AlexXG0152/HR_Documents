# HR_Documents
Filling HR documents for user from API data

ENGLISH VERSION. (RUSSIAN VERSION BELLOW)

I started this project to facilitate and unify paperwork in my division and at my enterprise (this is a large chemical plant). Every day we receive about 50-100 such documents. Since it takes time to fill them out, everyone tries to fill them in their own way, I (head of HR department) decided to make an application. It works through command line, user enters data gradually in correct order, receives a finished Word document and can save it on his PC, send it by mail, print, etc.
Application receives data from a server that contains a publicly available (in the local network) file with employee data. For the most part, the application accepts data in "str" format. I decided to make it easier for user to work with application. And Russian language is quite complicated in its rules, I could not foresee all variants of correct spelling and exceptions - user can make edits to finished document.
Server mentioned above is my API so to speak. It receives request and provides an HTML response. Application parses necessary table in the received HTML and takes information it needs for further work or displays information to user that requested data is not available on server.
Some data entered directly into the application (for example, the choice of the reason for dismissal). I did this for reason that only for these three reasons we dismiss employees without additional consultations on paperwork and the names of these reasons most likely will not change (established by labor legislation).

In total, application consists of 20 functions, each of which does its job and can be reused.
1. exception_handler() - catches errors and prevents application from breaking due to this.
2. resource_path() - used by PyInstaller to create one EXE file with application.
3. create_request_to_database() - creates data for API request.
4. send_request_to_db() - sends a request to API.
5. parsing_server_answer() - parses response from API.
6. create_list_of_holidays() - creates a list of holidays for calculating days of labor leave (according to the legislation, holidays are not vacation days, in this regard, it turns out +1 day for vacation). It is slightly different for each year due to religious holidays.
7. check_holidays() - checks if holidays days fall within holiday period.
8. date_input() - input dates for filling out forms (No. 2-7).
9. input_date_for_vacation() - input dates for calculating and filling out leave form (No. 1).
10. count_dates_for_vacation() - calculation of vacation dates depending on choice and user input.
11. format_date_for_template() - date formatting in accordance with rules of the Russian language.
12. format_date_for_template() - display employee's desire to receive annual financial assistance for part of vacation.
13. replace_wrong_cex_names() - API sometimes gives an incomplete long name for structural units and application replace it to right names.
14. data_from_csv_file() - receives data from CSV files (section in division, name of employer, types of liability.
15. initials() - abbreviation of the name and father name.
16. sklonenie() - declension and correct display of full name, position, profession and structural unit in accordance with rules of the Russian language.
17. dismissal_reason() - select the reason for dismissal.
18. data_for_liability_contract() - obtaining information to fill out an agreement on full individual liability.
19. fill_and_show_ready_template () - filling and displaying a DOC document selected by the user.
20. main () - logic of the entire application.

I would be very grateful for your constructive criticism.


RUSSIAN VERSION.

Я начал этот проект для облегчения и унификации оформления документов в своем подразделении и на своем предприятии (это большой химический завод). Ежедневно к нам поступает порядка 50-100 подобных документов. Так как на их заполнение требуется время, каждый пытается заполнить их по-своему, я решил сделать приложение. Работает через командную строку, пользователь вводит данные постепенно в правильном порядке, получает готовый документ Ворд и может сохранить его у себя на ПК, переслать по почте, распечатать и т.д.
Приложение получает данные от сервера, на котором находится общедоступная (в локальной сети) картотека с данными работников. По большей части данные приложение принимает в формате "str". Я решил сделать так, чтобы облегчить работу с приложением для пользователя. И т.к. русский язык довольно таки сложен в своих правилах, все варианты правильного написания и исключения я предусмотреть не смог - пользователь может внести правки в готовый документ.
Упомянутый выше сервер – это так сказать мое API. Он получает запрос и дает ответ в виде HTML. Приложение парсит необходимую таблицу в полученном HTML и берет в дальнейшую работу необходимую ему информацию. Либо отображает пользователю информацию о том, что запрошенных данных нет на сервере.
Некоторые данные внесены прямо в приложение (например, выбор причины увольнения). Я это сделал по той причине, что только по этим трем основаниям мы увольняем сотрудников без дополнительных консультаций по оформлению документов и названия этих причин скорее всего не поменяются (установлены трудовым законодательством).

Всего приложение состоит из 20 функций каждая из которых делает свою работу и может быть использована повторно.
1.  exception_handler() - ловит ошибки и не дает программе ломаться в связи с этим.
2. resource_path() - используется PyInstaller для создания одного EXE файла с программой.
3. create_request_to_database() - создает данные для API запроса.
4. send_request_to_db() - отправляет запрос к API.
5. parsing_server_answer() - парсит ответ от API.
6. create_list_of_holidays() - создает список праздников для расчета дней трудового отпуска (по законодательству праздничные дни не являются днями отпуска, в связи с этим получается +1 день к отпуску). Для каждого года он немного разный в связи с религиозными праздниками.
7. check_holidays() - проверяет не попадают ли дни в период  отпуска праздничные дни.
8. date_input() - ввод дат для заполнения бланков (№№2-7).
9. input_date_for_vacation() - ввод дат для расчета и заполнения бланка записки об отпуске (№1).
10. count_dates_for_vacation() - расчет дат отпуска в зависимости от выбора и ввода пользователя.
11. format_date_for_template() - форматирование даты в соответствии с правилами русского языка.
12. format_date_for_template() - для отображения желания работника получить ежегодную материальную помощь именно к этой части отпуска.
13. replace_wrong_cex_names() - API иногда отдает неполное длинное название структурных подразделений.
14. data_from_csv_file() - получает данные CSV файлов (участок в подразделении, ФИО нанимателя, виды материальной ответственности.
15. initials() - для сокращение имени и отчества.
16. sklonenie() - для склонения и правильного отображение ФИО, должности, профессии и структурного подразделения в соответствии с правилами русского языка.
17. dismissal_reason() - для выбора причины увольнения.
18. data_for_liability_contract() - получение информации для заполнения договора о полной индивидуальной материальной ответственности.
19. fill_and_show_ready_template() -заполнение и отображение выбранного пользователем DOC  документа.
20. main() - логика работы всего приложения.

Буду премного благодарен за конструктивную критику.
