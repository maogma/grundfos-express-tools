class User:
    def __init__(self, username, password):
        self.username = username
        self.password = password
    @classmethod
    def from_dict(cls, data):
        return cls(data['username'], data['password'])

users = [
    { 'username': 'rolf', 'password': '123' },
    { 'username': 'tecladoisawesome', 'password': 'youaretoo' }
]

# user_objects = map(User.from_dict, users)
