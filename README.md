# HTTPRequestScheduler
 **Goes row by row in an excel worksheet and create a http request for each row based on the data in it.**

## Using Pointers
The pointer is designed to pull text from the worksheet and can be used anywhere by placing it between two hash-symbols ('#').

There are two types of pointers: 
 * a static pointer that point to a specific cell and contains column letters + row digits. (e.g. #A6#)
 * a dynamic pointer that point to a cell on the current row and contains just column letters. (e.g. #A#)

Pointers can be placed on the worksheet itself by checking the "Recursive Insert" check box (or set it to "true" as argument).

![](https://github.com/KorenBar/HTTPRequestScheduler/blob/master/Captures/capture.gif)
