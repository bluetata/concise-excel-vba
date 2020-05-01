#### ByVal and ByRef in VB .NET

When you pass arguments over to Subs and Function you can do so either By Value or By Reference. By Value is shortened to **ByVal** and By Reference is shortened to **ByRef**. ByVal means that you are passing a copy of a variable to your Subroutine. You can make changes to the copy and the original will not be altered. Visual Studio hides ByVal from you most of the time. It's hidden because ByVal is the default when you're passing variables over to a function or Sub.

**ByRef** is the alternative. This is short for By Reference. This means that you are not handing over a copy of the original variable but pointing to the original variable. Let's see a coding example.

Add a new button the form you created in the previous section. Double click the button and add the following code:

```VB
Dim Number1 As Integer

Number1 = 10
Call IncrementVariable(Number1)

MessageBox.Show(Number1)
```

You'll get a wiggly line under **IncrementVariable(Number1)**. To get rid of it, add the following Subroutine to your code:


```VB
Private Sub IncrementVariable(ByVal Number1 As Integer)
Number1 = Number1 + 1
End Sub
```

When you're done, run the programme and click your new button. What answer was displayed in the message box?

It should have been 10. But hold on. Didn't we increment the variable Number1 with this line?

**Number1 = Number1 + 1**

So Number1 started out having a value of 10. After our Sub got called, we added 1 to Number1. So we should have 11 in the message box, right?

The reason Number1 didn't get incremented was because we specified **ByVal** in the Sub:

**ByVal Number1 As Integer**

This means that only a copy of the original variable got passed over. When we incremented the variable, only the copy got 1 added to it. The original stayed the same - 10.

Change the parameter to the this:

**ByRef Number1 As Integer**

Run your programme again. Click the button and see what happens.

This time, you should see 11 displayed in the message box. The variable has now been incremented!

It was incremented because we used ByRef. We're referencing the original variable. So when we add 1 to it, the original will change.

The default is **ByVal** - which means a copy of the original variable. If you need to refer to the original variable, use **ByRef**.



Ref link: https://www.homeandlearn.co.uk/NET/nets9p4.html
