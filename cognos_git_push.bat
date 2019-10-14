@echo off

setlocal EnableDelayedExpansion

SET "repoBase="
SET "qlikBase=C:\Program Files\ibm\cognos\analytics\deployment"
::SET "qlikFile=Sales.qvf"

SET "InvokeDesktop=YES"
SET qvfSrc="%qlikBase%%qlikFile%"

SET "gitPush=YES"

SET gitrepo="https://notijohn:qw3rty4321!@github.com/notijohn/cognos.git"

SET "applogs=event.log"
SET "qlikExe=C:\Users\amit.u.sharma\AppData\Local\Programs\Qlik\Sense\QlikSense.exe"
::SET "gitExe="
::SET "nodeExe=C:\Program Files\nodejs\node.exe"
SET "nodeCode=C:\Users\amit.u.sharma\Desktop\qliksense\Automation\QlikSenseDevOps"
SET "workenv=C:\Users\amit.u.sharma\Desktop\qliksense\workspace"

:: Dont edit below this 

::echo exec node command
::cd %nodeCode%
::git pull 
::"%nodeExe%" node_modules/qs-version-control-dev to-json -c config_demo.json

set mydate=!date:~10,4!!date:~6,3!/!date:~4,2!

echo commiting changes to git
xcopy "C:\Program Files\ibm\cognos\analytics\deployment\Test_db.zip" "C:\CognosDevOps\" /y
cd C:\CognosDevOps\
git remote add origin https://github.com/notijohn/cognos.git
git pull origin master
git add .
git commit -m "New files to upload"
git push --all https://github.com/notijohn/cognos.git
::git push --all %gitrepo%
set /p version_id=Enter version_id: 

:: Git Tagging

git tag %version_id%.%mydate%  
git push https://github.com/notijohn/cognos.git %version_id%.%mydate%   
::git push %gitrepo% %version_id%.%mydate%   
 
pause    

exit
