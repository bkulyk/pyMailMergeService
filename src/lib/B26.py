class B26:
    @staticmethod
    def fromBase26( value ):
        total = 0
        pos = 0
        for digit in value[::-1]:
            dec = int( digit, 36 ) - 9
            if pos == 0:
                add = dec
            else:
                x = dec * pos
                add = x * 26
            total = total + add
            pos = pos + 1
        return total