# excel_regex_clean
Golang application that performs a search & replace on Excel workbooks (except .xlsb extension), sheets within those workbooks, and cells within those sheets that match the regex patterns for each. regex expressions are specified in a .toml file with the same base name. 

### Prerequisites
```
"github.com/360EntSecGroup-Skylar/excelize"
"github.com/naoina/toml"
```
## Built With

* [Go 1.15.5]

## Running the Application
Run from the windows command window:
```
go run excel_replace_multiple.go
```

## Author
* **Bill Thaman** - *Initial work* - [bthaman/excel_regex_clean](https://github.com/bthaman/excel_regex_clean)