
from collections import UserList
from functools import wraps


def fill_until_index(f):
    @wraps(f)
    def decorated(self, index, *args, **kw):
        try:
            return f(self, index, *args, **kw)
        except IndexError:
            if self.default_factory is not None:
                for i in range(len(self.data), index + 1):
                    self.data.append(self.default_factory())
                return f(self, index, *args, **kw)
            else:
                raise
    return decorated


class DefaultList(UserList):

    def __init__(self, default_factory=None, data=None, **kw):
        super(DefaultList, self).__init__(data, **kw)

        self.default_factory = default_factory

    @fill_until_index
    def __getitem__(self, index):
        return self.data[index]

    @fill_until_index
    def __setitem__(self, index, item):
        self.data[index] = item
