import xlrd
from xlrd.timemachine import xrange

# src = 'wonders.xls'
src = 'Wonders_Reading_Writing_Workshop.xlsx'
with xlrd.open_workbook(src) as book:
    # 0 corresponds for 1st worksheet, usually named 'Book1'
    sheet = book.sheet_by_index(3)

    # gets col D values
    D = [D for D in sheet.col_values(3)]

    # gets col E values
    E = [E for E in sheet.col_values(4)]

    # combines D and E elements to tuples, combines tuples to list
    # ex. [ ('Incoming', 18), ('Outgoing', 99), ... ]
    data = list(zip(D, E))

    # gets sum
    # incoming_sum = sum(tup[1] for tup in data if tup[0] == 'Incoming' )
    # outgoing_sum = sum(tup[1] for tup in data if tup[0] == 'Outgoing' )

    # print('Total incoming:', incoming_sum)
    # print('Total outgoing:', outgoing_sum)

    print(data)


    class Cell(object):
        def __init__(self, r, c):
            self.row = r
            self.col = c
            pass


    class MergedCell(Cell):
        def __init__(self, t):
            rlo, rhi, clo, chi = t
            Cell.__init__(self, rlo, clo)
            self.row_h = rhi
            self.col_h = chi

        pass


    class Item(object):
        def __init__(self, name, title):
            self.id = name
            self.title = title

        pass


    class Course(Item):
        def __init__(self, name, title, question, reading, c_strategy, c_skill, genre, v_strategy, writing):
            Item.__init__(self, name, title)
            self.question = question
            self.reading = reading
            self.c_strategy = c_strategy
            self.c_skill = c_skill
            self.genre = genre
            self.v_strategy = v_strategy
            self.writing = writing

        pass


    class Unit:
        def __init__(self, name, title, idea):
            Item.__init__(self, name, title)
            self.idea = idea
            self.courses = []

        pass


    units = {}

    # sort merged_cells by its low row first and its low column later
    sorted_merged_cells = sorted(sheet.merged_cells, key=lambda t: sheet.utter_max_cols * t[0] + t[2])

    # traverse sorted merged cells and parse units dictionary
    name = None
    title = None
    idea = None
    row = None
    for crange in sorted_merged_cells:
        rlo, rhi, clo, chi = crange
        cell = sheet.cell(rlo, clo)
        if clo == 0:
            name = cell
        elif clo == 1:
            title = cell
        elif clo == 2:
            idea = cell

        if row is None:
            row = rlo

        if rlo != row:
            row = rlo
            units[row] = Unit(name, title, idea)
            pass

        # print((rlo, clo), cell)
        for rowx in xrange(rlo, rhi):
            for colx in xrange(clo, chi):
                # print(sheet.cell(rlo, clo), sheet.cell(rowx, colx))
                pass
        # cell (rlo, clo) (the top left one) will carry the data
        # and formatting info; the remainder will be recorded as
        # blank cells, but a renderer will apply the formatting info
        # for the top left cell (e.g. border, pattern) to all cells in
        # the range.

    print(len(units))
    for i in units:
        u = units[i]
        print(u.idea)
