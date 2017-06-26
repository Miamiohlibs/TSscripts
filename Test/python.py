from time import sleep

#
# greeting = 'hello'
# name = input("Please enter our name")
# print(greeting + " " + name)

# splitString = "this string has been \nsplit over\nlines"
# print(splitString)
# sleep(3)
#
# anotherSplitString = """This string has
# been split
# over three lines"""
# print(anotherSplitString)


# a = 12
# b = 3
# c = a + b
# d = c / 3
# e = d - 4
# print(e*12)
#
# parrot = "norwegian blue"
# print(parrot)
# print(parrot[-3])
# print(parrot[0:6])
#
# today = "friday"
# print("day" in today)
# print("parrot" in today)


# age = 24
# print("my age is " + str(age) + " years")
#
# print("My age is {0} years".format(age))
# print("there are {0} days in {1}, {2}, {3}, {4}, {5}, {6}, {7}".format(31, "january", "mar", "may","jul","aug","oct","dec"))
#

# print("""Jan: {2}
# feb: {0}
# mar: {2}
# apr: {1}
# may: {2}
# jun: {1}
# jul: {2}
# aug: {2}
# sep: {1}
# oct: {2}
# nov: {1}
# dec: {2}""".format(28,30,31))
#

# for i in range(1,12):
#     print("no. %2d squared is %4d and cubed is %4d" %(i, i**2, i**3))

#
# name = input("please enter your name: ")
# age = int(input("how old are you, {0}? ".format(name)))
# print(age)
# if age >=18:
#     print("you are old enough to vote")
#     print("please put an X in the box")
# else:
#     print("please come back in {0} years".format(18-age))


# print("please guess a number between 1 and 10: ")
# guess = int(input())
#
# if guess != 5:
#     if guess <5:
#         print("please guess higher")
#     else:
#         print("please guess lower")
#     guess = int(input())
#     if guess == 5:
#         print("well done, you guessed it")
#     else:
#         print("sorry you have not guessed correctly.")
#
# else:
#     print("you got it first time")


# age = int(input("How old are you? "))
# #if age >= 16 and age <=65:
# # if 16 <= age <= 65:
# #     print("have a good day at work")
# if (age < 16) or (age>65):
#     print("enjoy your free time")
# else:
#     print("have a good day at work")

age = int(input("how old are you? "))
if not(age<18):
    print("old enough")
    print("x in box")
else:
    print("please come back in {0}".format(18-age))