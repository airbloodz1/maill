
import win32com.client as client

# создаем объект Outlook
outlook = client.Dispatch('Outlook.Application')

# получаем объект Namespace
namespace = outlook.GetNamespace('MAPI')
# создаем объект письма
mail = outlook.CreateItem(0)

# Сохраняем изображение
attachment = mail.Attachments.Add(r'C:\Users\s.kondratyuk\Pictures\подписьFozzy.jpg', 5, 0, 'MyImage')

# Получаем полный путь к сохраненному изображению
image_path = attachment.PathName

#  HelpUkranians.jpg

# Создаем HTML-код для подписи
signature = '<p>СЕРГІЙ<br>КОНДРАТЮК<br>Аналітик комп’ютерного банку даних<br>Департамент звітності<br>+380 (99) 923-63-41</p>'
signature += f'<p><img src="file://{image_path}"></p>'



# устанавливаем параметры письма
mail.To = 's.kondratyuk@fozzy.ua'
mail.Subject = 'Тема письма'
mail.Body = 'Текст письма'


mail.HTMLBody = f"{mail.HTMLBody}{signature}"

# отправляем письмо
mail.Send()