# May
May is an Excel VBA tool and class, it contains many usefull and powerful VBA functions, these functions are not independent, they are a total system. 
If need any help, pls send mail to  shezhidian@163.com, i will try my best to help VBA developers.
You can download MayManual.xlsm to find all function details.

Quick Start:


Step1: Import MayClass.cls、MayClassPosition.cls into VBA project


<img width="112" alt="image" src="https://user-images.githubusercontent.com/69334389/218962665-65dc3b37-8fa1-4d34-9bac-de51000928e1.png">


Step2: Hello World!


Sub testHelloWord()

  Dim may As New MayClass

  Debug.Print may.AboutMay

End Sub



Step3: Test One import function: ArrayColumnTakeByTitle()


first: Prepare test data such like:


![image](https://user-images.githubusercontent.com/69334389/218956399-adf55467-14c6-4d08-9513-e1dba6269341.png)

then coding:

Sub testArrayColumnTake()


    Dim may As New MayClass
    
    Dim data As Variant
    
    Dim res As Variant
    
    'selection is the data you prepared at first
    
    data = Selection.Value
    
    res = may.ArrayColumnTakeByTitle(data, "t2", True)
    
    may.ShowData res


End Sub

run results:


![image](https://user-images.githubusercontent.com/69334389/218956800-ae336855-fa5a-4783-aab6-4160b325c9d4.png)
