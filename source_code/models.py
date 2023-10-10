class People:
    name = ""
    status = ""
    free_time = []
    sche_time = []
    compensation = 0

    def __init__(self, name, status):
        self.name = name
        self.status = status
        self.free_time = []
        self.sche_time = []
        self.compensation = 0
