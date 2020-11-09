class MyException(Exception):
    """
    自定义异常，用来测试自己的代码。
    """

    def __init__(self, error):
        self.error = error

    def __str__(self, *args, **kwargs):
        return self.error
