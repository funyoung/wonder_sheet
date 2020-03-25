import xlrd
from xlrd.timemachine import xrange

src = 'wonders.xls'
with xlrd.open_workbook(src, formatting_info=True) as book:
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

    for crange in sheet.merged_cells:
        rlo, rhi, clo, chi = crange
        for rowx in xrange(rlo, rhi):
            for colx in xrange(clo, chi):
                print(sheet.cell(rlo, clo), sheet.cell(rowx, colx))
        # cell (rlo, clo) (the top left one) will carry the data
        # and formatting info; the remainder will be recorded as
        # blank cells, but a renderer will apply the formatting info
        # for the top left cell (e.g. border, pattern) to all cells in
        # the range.
