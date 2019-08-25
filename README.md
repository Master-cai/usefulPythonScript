# Project Title excelKiller

The project is used to deal with excel merging jobs in some specific case

## Getting Started 

### Prerequisites



​	The project run on python3.7 with anconda, a package management tool.

​	To use it, you need to make sure the excels` file name is same as the name in the excel file.

​	First at all, you need to clone it. Just put the python file in the same dictionary with your excels. 

then you should make a text file named `nameList.txt`  to save all the people`s name you want to check or to be merged. The names should write in a line and Separate by a full width comma. And you need to create a excel file which you want to keep the output. I recommend you that you can name it  output.xls which is default.

​	After all, just run in a python environment.

`python excelKiller.py`

​	The command without any parameter will use the default parameter and the result will be saved in the  dictionary which you clone the Project.

​	The result may not be what you want. Don`t worry, you can use the command parameter to make it be suitable to your job.

### feature

​	According to the nameList.txt provided by you, it will print the file has been merged and the file didn\`t find which means it does\`t exist !

#### Optional parameter

- '-w' or '--workPath': the param is to set the path of your excels` dictionary, default is current working directory.
- '-o' or '--oldRowNum': the param is to set the row number you want to insert according to your own case.
- '-c' or '--column': the param is to set the column number of the name column.You should know that it begins with 0.
- '-O' or '--outPutFiles':the param is to set the file you want ro save the output, it must exist.
- '--help': For all the parameter above, you can use it to find the prompt message.

### Usage example

use with your own parameter:

`python excelKiller.py -w D:\OneDrive -o 2 -O D:\OneDriveoutput.xls -c 3`

to get help message, you should it:

`python excelKiller.py --help`



## Contributing

Please read [CONTRIBUTING.md](https://gist.github.com/wangyan/6e8021667fe7f2082d153bed2d764618#) for details on our code of conduct, and the process for submitting pull requests to us.

## Release History

- 0.0.1 

## Authors

GitHub:[Master-cai](https://github.com/Master-cai)

## License

For this project MIT agreement, please click LICENSE.md for more details.