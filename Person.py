class Person(object):
    '员工的基本类'

    def __init__(self, secDepartment, name, tbName, bianzhi, company, newPrice, oldPrice, thirdDepartment, position,
                 email, phone, enterTime, outTime):
        self.secDepartment = secDepartment
        self.name = name
        self.tbName = tbName
        self.bianzhi = bianzhi
        self.company = company
        self.newPrice = newPrice
        self.oldPrice = oldPrice
        self.thirdDepartment = thirdDepartment
        self.position = position
        self.email = email
        self.phone = phone
        self.enterTime = enterTime
        self.outTime = outTime
