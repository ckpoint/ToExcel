
# To-Excel

This is a library that makes Excel easier to use in Java.

To-Excel depends on  [Apache POI](https://poi.apache.org), [Model Mapper](https://github.com/modelmapper/modelmapper), [Project Lombok](http://projectlombok.org/)


# Installation

#### 1. MAVEN
```xml
<dependency>
  <groupId>com.github.ckpoint</groupId>
  <artifactId>toexcel</artifactId>
  <version>0.2</version>
</dependency>

```
#### 2. GRADLE
```gradle
  compile group: 'com.github.ckpoint', name: 'toexcel', version: '0.2'
```


# Usage

## Table of Contents
- [ 1. Map To Object From Excel ](#map-to-object-from-excel)
- [ 2. Map To Excel From Object ](#map-to-excel-from-object)
- [ 3. Create Custom Excel Sheet](#create-excel-sheet)

## 1. Map to Object from Excel
#### Read data from a specific sheet in Excel and can map data to a specific class.

### step1) Create a model class to map Excel data

```
name | age | gender
Ace  | 20  | MAN
Eve  | 23  | MAN
Eva  | 20  | WOMAN
        .
        .
        .
```

#### If you want to read the above Excel data, you need to define the model as below.


```java
public class UserModel {

    @ExcelHeader(headerName = "name", headerNames = {"nick-name, nickName, email"})
    private String name;

    @ExcelHeader(headerName = "age")
    private Integer age;

    @ExcelHeader(headerName = "gender")
    private String gender;
}
```
#### You need to add the ExcelHeader annotation to each field.
##### headerName must match the data header of Excel.
##### If there is no data field in Excel that matches the headerName, can search for the matching field in headerNames.

### step2) Create ToWorkBook and ToWorkSheet instance

```java
        ToWorkBook toWorkBook = new ToWorkBook("target/excel/map/read_test_1.xlsx");
        ToWorkSheet toWorkSheet = toWorkBook.getSheetAt(0);
```

### step3) You can now map data to model classes using ToWorkSheet's map function.

```java
        List<UserModel> userModels = toWorkSheet.map(UserModel.class);
```

## 2. Map To Excel From Object
#### You can easily map data defined in the model to Excel.

#### step1) define the model class to be mapped to Excel as below.

```java
@Builder
public class UserModel {

    @ExcelHeader(headerName = "name", priority = 0)
    private String name;
    @ExcelHeader(headerName = "age", priority = 1)
    private Integer age;
    @ExcelHeader(headerName = "gender", priority = 2)
    private String gender;
}
```

#### You need to add the ExcelHeader annotation to each field.
##### headerName must match the data header of Excel.
##### The order in which Excel is written, the lower the priority, the write left.

### step2) Define data in the model
```java
 List<UserModel> userModelList =
         IntStream.range(0, 100).mapToObj(i ->
         UserModel.builder().name("tester" + i).age(i).gender("man").build()).collect(Collectors.toList());

```

### step3) Create instance ToWorkBook and ToWorkSheet
```java
        ToWorkBook workBook = new ToWorkBook(WorkBookType.XSSF);
        ToWorkSheet sheet = workBook.createSheet();
```

### step4) Use ToWorkSheet's from function to map data to a sheet and write it to a file.

```java
   sheet.from(userModelList);
   workBook.writeFile("target/excel/map/write_test_1");
```

## 3. Create Custom Excel Sheet
#### You can freely write data to Excel without defining a model.

#### For example, you can write the following like code
```java
    ToWorkBook workBook = new ToWorkBook(WorkBookType.XSSF);
    ToWorkSheet sheet = workBook.createSheet();

        sheet.createTitleCell(2, "name", "age");
        sheet.next();
        IntStream.range(0, 50).forEach(i ->{
            sheet.createCell("Jack" + i , 1 + i);
            sheet.next();
        });

    workBook.writeFile("target/excel/test.xls");
```
