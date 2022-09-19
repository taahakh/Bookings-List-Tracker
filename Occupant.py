from datetime import datetime
from settings import DATE_FORMAT, DATE_FORMAT2

class Occupant():

    cleaned_end = ""

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

    def equals(self, occupant) -> bool:
        if self.address.__eq__(occupant.address) and self.room == occupant.room and self.name.__eq__(occupant.name) and self.ref == occupant.ref and self.room_size.__eq__(occupant.room_size) and self.start_date.__eq_(occupant.start_date) and self.end_date.__eq_(occupant.end_date) and self.rate.__eq_(occupant.rate) and self.number_of_nights.__eq_(occupant.number_of_nights):
            return True
        return False

    def end_occupancy(self) -> bool:
        if self.end_date.value is None:
            return False

        cleaned_date = str(self.end_date.value).rstrip()

        # print(cleaned_date)

        if len(cleaned_date) != 0:
            try:
                self.cleaned_end = datetime.strptime(cleaned_date, DATE_FORMAT)
                # self.cleaned_end = str(datetime.strptime(cleaned_date, DATE_FORMAT))
                # print("TUPE1: ", self.cleaned_end)
            except:
                try:
                    self.cleaned_end = datetime.strptime(cleaned_date, DATE_FORMAT2)
                    # self.cleaned_end = str(datetime.strptime(cleaned_date, DATE_FORMAT2))
                    # print("TUPE2: ", self.cleaned_end)
                    return True
                except:
                    pass
                return False
            return True

        return False

    def correct_invoice(self, invoice):
        # occupant and placement
        # NEED TO DO REF CHECK
        # if self.name.value != invoice.name.value or int(self.ref.value) != int(invoice.ref.value):
        # if self.name.value != invoice.name.value or str(int(self.ref.value)) != str(invoice.ref.value):
        if self.name.value.lstrip().rstrip() != invoice.name.value.lstrip().rstrip():
            return False

        # print("pass1")
        # address
        # not all address names that are meant to be the same are the same e.g road is rd in other document
        if self.address.value.lstrip().rstrip() not in invoice.address.value.lstrip().rstrip():
            return False

        # print("pass2")
        # room no
        if int(self.room.value) != int(invoice.room.value):
            return False

        # print("pass3")
        # room size
        if self.room_size.value.lstrip().rstrip() not in invoice.room_size.value.lstrip().rstrip():
            return False

        # print("pass4")
        # end date
        # STANDARDISE END DATE
        # if self.cleaned_end != invoice.cleaned_end:
        #     return False

        # nightly rate
        # some nightly rates are missing
        # when they are updated on the tracker, the code isn't going to pick up there is a
        # current invoice already there. Reduce correct criteria to name, address, room no, room size
        # there needs to be a function that searches for missing values and updates them when needed
        if int(float(self.rate.value[1:])) != int(invoice.rate.value):
            return False

        # print("pass5")
        return True

    # true when the month is old now and needs to be deleted
    # false when the month is the same. Need to update the current month instead
    def need_to_delete_invoice(self, current_month) -> bool:
        if self.cleaned_end.month == current_month:
            return False
        return True

    def compare_address(self, address) -> bool:
        if str(self.address.value).lstrip().rstrip() not in str(address).lstrip().rstrip():
            return False

        return True

    def compare_name(self, name) -> bool:
        if str(self.name.value).lstrip().rstrip() != str(name).lstrip().rstrip():
            return False

        return True

    def compare_address_name(self, name, address) -> bool:
        if str(self.address.value).lstrip().rstrip() not in str(address).lstrip().rstrip():
            return False

        if str(self.name.value).lstrip().rstrip() != str(name).lstrip().rstrip():
            return False

        return True