## Как собрать и запустить проект в docker образе локально
Запускать из корневой папки с солюшен файлом.
```
docker build -f .\Microsoft.Exchange.WebServices.Data.Test\Dockerfile -t exchange-tests .
docker run -it exchange-tests -e EwsServiceUrl='' -e Username='' -e Password=''
```