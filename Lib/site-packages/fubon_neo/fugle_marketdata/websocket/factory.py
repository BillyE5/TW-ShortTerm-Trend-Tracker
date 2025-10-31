from ..enums import Mode
from ..client_factory import ClientFactory
from .futopt import WebSocketFutOptClient
from .stock import WebSocketStockClient
from ..constants import FUGLE_MARKETDATA_API_NORMAL_WEBSOCKET_BASE_URL, FUGLE_MARKETDATA_API_NORMAL_VERSION, FUGLE_MARKETDATA_API_SPEED_WEBSOCKET_BASE_URL, FUGLE_MARKETDATA_API_SPEED_VERSION

class WebSocketClientFactory(ClientFactory):
    def __init__(self, mode = Mode.Speed, **options):
        super().__init__(**options)
        self.__clients = {}
        self.mode = mode
        self.options = options

    @property
    def stock(self):
        return self.get_client('stock')

    @property
    def futopt(self):
        return self.get_client('futopt')

    def get_client(self, type):

        base_url = ''
        if self.mode == Mode.Normal:
            base_url = f"{FUGLE_MARKETDATA_API_NORMAL_WEBSOCKET_BASE_URL}/{FUGLE_MARKETDATA_API_NORMAL_VERSION}/{type}/streaming"
        elif self.mode == Mode.Speed:
            base_url = f"{FUGLE_MARKETDATA_API_SPEED_WEBSOCKET_BASE_URL}/{FUGLE_MARKETDATA_API_SPEED_VERSION}/{type}/streaming"

        if type in self.__clients:
            return self.__clients[type]
        
        if type == 'stock': 
            client = WebSocketStockClient(self.mode, base_url=base_url, **self.options)
        elif type == 'futopt' :
            client = WebSocketFutOptClient(self.mode, base_url=base_url, **self.options)
        else: 
            None

        self.__clients[type] = client
        return client





