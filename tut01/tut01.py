def factorial(x):
    factorial = 1
    if x < 0 :
        return ('factorial does no exist for negative numbers')
    elif x ==0 :
        return 1
    else :
        for i in range (1,x+1) :
            factorial = factorial*i
        return factorial

x=int(input("Enter the number whose factorial is to be found : "))
print(factorial(x))