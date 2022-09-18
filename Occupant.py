from datetime import datetime

DATE_FORMAT = "%d/%m/%Y"
DATE_FORMAT2 = "%Y-%d-%m 00:00:00"

class Occupant():

    address = ""
    room = 0
    name = ""
    ref = 0
    room_size = ""
    start_date = ""
    end_date = ""
    rate = ""
    number_of_nights = 0

    def __init__(self, address, room, name, ref, room_size, start_date, end_date, rate, number_of_nights=31):
        self.address = address
        self.room = room
        self.name = name
        self.ref = ref
        self.room_size = room_size
        self.start_date = start_date
        self.end_date = end_date
        self.rate = rate
        self.number_of_nights = number_of_nights

    # def __init__(self, name, ref):
    #     self.name = name
    #     self.ref = ref

    def equals(self, occupant) -> bool:
        if self.address.__eq__(occupant.address) and self.room == occupant.room and self.name.__eq__(occupant.name) and self.ref == occupant.ref and self.room_size.__eq__(occupant.room_size) and self.start_date.__eq_(occupant.start_date) and self.end_date.__eq_(occupant.end_date) and self.rate.__eq_(occupant.rate) and self.number_of_nights.__eq_(occupant.number_of_nights):
            return True
        return False

    # REMOVE WHITESPACE AS WELLLLLLLLLLL
    def end_occupancy(self) -> bool:
        # if self.ref.value == 239125.0:
        #     print(self.end_date.value)
        #     print("ok")
        # print(self.end_date.value)
        if self.end_date.value is None:
            return False

        if len(str(self.end_date.value)) != 0:
            try:
                if self.ref.value == 239125.0:
                    print(str(self.end_date.value))
                    print("ok")
                # print(datetime.strptime(self.end_date.value, DATE_FORMAT))
                # print(str(self.end_date.value))
                datetime.strptime(str(self.end_date.value), DATE_FORMAT)
                # return True
            except Exception as e:
                try:
                    if self.ref.value == 239125.0:
                        print(e)
                        print(str(self.end_date.value))
                        print("ok2")
                    datetime.strptime(str(self.end_date.value), DATE_FORMAT2)
                    return True
                except:
                    return False
                # print("didnt work")
                return False
            return True

        return False