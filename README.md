![Go](https://github.com/sergrom/csv2xls/actions/workflows/go.yml/badge.svg)

# csv2xls
The fast <strike>and furious</strike> converter excel format csv into xls.

## Install
```bash
$ go install github.com/sergrom/csv2xls/v3@v3.0.1
```

## Usage
To see paramters and options type:
```bash
$ csv2xls -h
```
Example command for converting file:
```bash
$ csv2xls -csv-file-name="file.csv" -xls-file-name="file.xls"
```

## Explaination parameters and options
<code>--csv-file-name</code> - The csv file you want to convert. Mandatory parameter.<br>
<code>--xls-file-name</code> - The xls file name that will be created. Mandatory parameter.<br>
<code>--csv-delimiter</code> - The delimiter that used in csv file. Optional parameter. Default value is semicolon - ";".<br>
<code>--title</code> - The Title property of xls file. Optional parameter.<br>
<code>--subject</code> - The Subject property of xls file. Optional parameter.<br>
<code>--creator</code> - The Creator property of xls file. Optional parameter.<br>
<code>--keywords</code> - The Keywords property of xls file. Optional parameter.<br>
<code>--description</code> - The Description property of xls file. Optional parameter.<br>
<code>--last-modified-by</code> - The LastModifiedBy property of xls file. Optional parameter.

## Example
For example you have csv file with name <b>cities.csv</b> and you want to convert it into xls excel format. The content of csv file is, for example:
<pre>
№,City,Population
1,Moscow,"12,537,954"
2,"St Petersburg","5,467,808"
3,"Krasnoyarsk","1,137,494"
</pre>

You can convert it by running command in your terminal:
```bash
./csv2xls -csv-file-name="cities.csv" -xls-file-name="cities.xls" -csv-delimiter=","
```

And then you will have a newly created file <b>cities.xls</b> in the same directory.<br>
![xls](https://user-images.githubusercontent.com/17692545/75096799-20252180-55b4-11ea-8ffc-6986086f5163.png)
<br>
Enjoy)
