string = '12abcd405'
def removeDigits(string):
    numbers = ['0','1','2','3','4','5','6','7','8','9']
    result=""
    for i in string:
        if i not in numbers:
            result += i

    print(result)

removeDigits(string)