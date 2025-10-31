from fubon_neo._fubon_neo import (
    CoreSDK,
    Order,
    Condition,
    ConditionDayTrade,
    ConditionOrder,
    FutOptOrder,
    FutOptConditionOrder,
)
from .fugle_marketdata import WebSocketClient, RestClient, Mode

class MarketData:
    def __init__(self, sdk_token, mode):
        self.websocket_client =  WebSocketClient(mode = mode, sdk_token = sdk_token)
        self.rest_client = RestClient(sdk_token = sdk_token)

class FubonSDK(CoreSDK):
    """
    fubon sdk for api trading
    """

    def init_realtime(self, mode = Mode.Speed):
        """
        Initial market data and get authorised

        """
        sdk_token = super().exchange_realtime_token()
        self.marketdata = MarketData(sdk_token, mode)
