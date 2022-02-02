"""
Defines class Filters
"""
class Filters:
    """
    Auxilliary class to assist with Excel AutoFilter object
    """
    def __init__(self, table) -> None:
        self.table = table
        # self.previous_filters = []

    def show_all_data(self) -> None:
        """
        Show all data disregarding existing filters
        """
        if self.table.api.AutoFilter is not None:
            for i in range(self.table.api.ListColumns.Count):
                # print(filters.Filters(i).Criteria1)
                pass
            self.table.api.AutoFilter.ShowAllData()

    def restore_filters(self) -> None:
        """
        Restore any filters prior showing all
        """
        pass
