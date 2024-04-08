## Телеграм-бот для записи к врачу

## The Telegram bot for making an appointment with a doctor

Клиент нажимает кнопку "Записаться", выбирает в инлайновых кнопок врача, из следующего меню с кнопками - дату, и далее, по тому же принципу - время. Инлайновые кнопки пропадают после нажатия на них. 
Запись происходит в Excel-файл на сервере. Профессия врача берется из имени файла, дата - из имени листа, время - из первого столбца на листе.
Если ячейки справа свободны - клиент видит свободное время на этот день. Он пишет своё ФИО и номер телефона, данные записываются в таблицу, а клиент видит подтверждение в сообщении.

Python 3.10.8

![Phone](https://github.com/Demston/Family_Doctor_test_bot/blob/main/med%20bot%20screenshot.png)
