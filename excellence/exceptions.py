

class ExcellenceException(Exception):
    pass


class InvalidGrid(ExcellenceException):
    pass


class ImmutableCellError(ExcellenceException):
    pass
