from datetime import datetime

from openpyxl import load_workbook
from openpyxl.cell.read_only import EmptyCell
from openpyxl.styles import PatternFill, colors

red_background = PatternFill("solid", start_color=colors.RED, end_color=colors.RED)
green_background = PatternFill("solid", start_color=colors.GREEN, end_color=colors.GREEN)


class Reader:

    def __init__(self, inventory_filename: str, replenishment_filename: str):
        self.inventory_book = load_workbook(inventory_filename, read_only=True)
        self.replenishment_book = load_workbook(replenishment_filename)
        self.inventory_sheet = self.inventory_book.active
        self.replenishment_sheet = self.replenishment_book.active
        self.quantity_on_hand_column_number, self.inventory_sku_columns = self.find_inventory_columns()
        self.replenishment_sku_column_number = self.find_column_number(self.replenishment_sheet, 'sku')

    def find_inventory_columns(self):
        sku_columns = []
        headers = next(self.inventory_sheet.iter_rows(min_row=1, max_row=1))
        quantity_on_hand_column_number = None
        for column in headers:
            value = column.value
            if value is None or not column.data_type == 's':
                continue
            else:
                if 'sku' in value.lower():
                    sku_columns.append(column.column - 1)
                if 'hand' in value.lower():
                    quantity_on_hand_column_number = column.column - 1
        return quantity_on_hand_column_number, set(sku_columns)

    @staticmethod
    def find_column_number(sheet, column_name):
        for head in sheet.iter_rows(1, 1):
            for column in head:
                if column_name in column.value.lower():
                    return column.column

    def find_sku_quantity(self, sku):
        pass

    def run(self):
        print("building Inventory...")
        inventory = {}
        columns_itr = next(self.replenishment_sheet.iter_cols(self.replenishment_sku_column_number, min_row=2))
        data = {cell.value: cell.row for cell in columns_itr}
        for row in self.inventory_sheet.iter_rows(min_row=2):
            if isinstance(row, EmptyCell):
                continue
            columns = [row[x] for x in self.inventory_sku_columns if not isinstance(row[x], EmptyCell)]
            inventory.update({column.value: {'quantity': row[self.quantity_on_hand_column_number].value} for column in
                              columns if column.value in data})

        print("filling in colors...")
        for sku, item_data in inventory.items():
            row_index = data[sku]
            quantity = int(item_data['quantity'])
            row = self.replenishment_sheet[f"{row_index}:{row_index}"]
            for column in row:
                column.fill = green_background if quantity > 0 else red_background
        self.replenishment_book.save('updated_replenishment.xlsx')


if __name__ == '__main__':
    reader = Reader('inventory.xlsx', 'replenishment.xlsx')
    start = datetime.utcnow()
    # print(reader.find_sku('IX-BHSV-2ZPK'))
    reader.run()
    end = datetime.utcnow()
    print(f"Took: {(end - start).total_seconds():,} seconds")
