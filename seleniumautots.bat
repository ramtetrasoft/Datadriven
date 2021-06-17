C:
set JAVA_OPTS=%JAVA_OPTS% -Xms512m -Xmx2048m -XX:PermSize=1024m -XX:MaxPermSize=2048m
set projectLocation=C:\Users\tsipl1805\seleniumjava\org.ts.seleniumauto
cd %projectLocation%
set classpath=%projectLocation%\bin
set classpath=%classpath%;%projectLocation%\lib\*
java org.testng.TestNG\TestSuite\LoginTest.xml
pause