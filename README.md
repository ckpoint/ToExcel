
# To-Excel

This is a library that makes Excel easier to use in Java.

To-Excel depends on  [Apache POI](https://poi.apache.org), [Model Mapper](https://github.com/modelmapper/modelmapper)


# Installation

#### 1. MAVEN
```xml
<dependency>
  <groupId>com.github.ckpoint</groupId>
  <artifactId>toexcel</artifactId>
  <version>1.0.3</version>
</dependency>

```
#### 2. GRADLE
```gradle
  compile group: 'com.github.ckpoint', name: 'toexcel', version: '1.0.3'
```


# Usage

## Table of Contents
- [ 1. Map To Object From Excel ](#1-map-to-object-from-excel)
- [ 2. Map To Excel From Object ](#2-map-to-excel-from-object)
- [ 3. Create Custom Excel Sheet](#3-create-custom-excel-sheet)

## 1. Map to Object from Excel
#### Read data from a specific sheet in Excel and can map data to a specific class.

### step1) Create a model class to map Excel data

![image](https://user-images.githubusercontent.com/30170928/66096579-657d8e00-e5d6-11e9-85af-c39dec335ece.png)

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
   ToWorkBook toWorkBook = new ToWorkBook(new File("target/excel/map/read_test_1.xlsx"));
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

##### Finally, the following Excel is created.

![image](https://user-images.githubusercontent.com/30170928/66096579-657d8e00-e5d6-11e9-85af-c39dec335ece.png)

## 3. Create Custom Excel Sheet
####  You can more freely write data to Excel without defining a model.

#### You can create the following sheets.

### A. cells merged in the horizontal direction

![image](https://user-images.githubusercontent.com/30170928/66097565-cb1f4980-e5d9-11e9-8f97-82e879620266.png)

```java
   ToWorkBook workBook = new ToWorkBook(WorkBookType.XSSF);
   ToWorkSheet sheet = workBook.createSheet().updateDirection(SheetDirection.HORIZON);

   sheet.createTitleCell(2, "name", "age", "contactNumber");
   sheet.merge(2, 1); // 2(width) X 1(height) [][]
   sheet.createCellToNewline("sharky", "36", "010-1234-0000", "02-1111-1234");
   sheet.createCellToNewline("melpis", "36", "010-1111-1234", "02-4221-1234");
   sheet.createCellToNewline("heeseob", "32", "010-0000-1234", "-");
   sheet.createCellToNewline("dongjun", "31", "010-4324-1234", "031-4121-1234");

   workBook.writeFile("target/excel/manual/merge/merge_horizon_01");
```

### B. Multiple cells merged horizontally and vertically

![image](https://user-images.githubusercontent.com/30170928/66097671-4ed93600-e5da-11e9-9e46-41eb898509be.png)

```java
   ToWorkBook workBook = new ToWorkBook(WorkBookType.XSSF);
   ToWorkSheet sheet = workBook.createSheet().updateDirection(SheetDirection.HORIZON);

   sheet.createTitleCell(2, "name");
   sheet.merge(1, 2);

   sheet.createTitleCell(2, "age");
   sheet.merge(1, 2);

   sheet.createTitleCell(2, "contactNumber");
   sheet.merge(2, 1);
   sheet.newLine();

   sheet.createTitleCell(2,"phone", "home");

   sheet.createCellToNewline("sharky", "36", "010-1234-0000", "02-1111-1234");
   sheet.createCellToNewline("melpis", "36", "010-1111-1234", "02-4221-1234");
   sheet.createCellToNewline("heeseob", "32", "010-0000-1234", "-");
   sheet.createCellToNewline("dongjun", "31", "010-4324-1234", "031-4121-1234");

   workBook.writeFile("target/excel/manual/merge/merge_horizon_02");
```


Thank you to Dami Im, who gave a lot of inspiration and help to the development of the source.

