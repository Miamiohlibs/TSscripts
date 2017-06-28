name = input("what is your name? ")
age = int(input("how old are you {0}? ".format(name)))
if (18 <= age < 31):
    print("welcome to holiday")
else:
    print("you are not 18-30")