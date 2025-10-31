import json
import websocket
from enum import Enum
from pyee import EventEmitter
from threading import Thread, Timer
from ..enums import Mode

from ..constants import (
    AUTHENTICATION_TIMEOUT_MESSAGE,
    CONNECT_EVENT,
    DISCONNECT_EVENT,
    MESSAGE_EVENT,
    ERROR_EVENT,
    AUTHENTICATED_EVENT,
    MISSING_CREDENTIALS_MESSAGE,
    UNAUTHENTICATED_EVENT,
    UNAUTHENTICATED_MESSAGE
)

websocket.setdefaulttimeout(5)


class AuthenticationState(Enum):
    PENDING = 0
    AUTHENTICATING = 1
    AUTHENTICATED = 2
    UNAUTHENTICATED = 3


class WebSocketClient():
    def __init__(self, mode, **config):
        self.mode = mode
        self.config = config
        self.ee = EventEmitter()
        self.ee.on(CONNECT_EVENT, self.__handle_authentication)
        self.__ws = websocket.WebSocketApp(
            self.config.get('base_url'),
            on_open=self.__on_open,
            on_close=self.__on_close,
            on_error=self.__on_error,
            on_message=self.__on_message)

        self.auth_timer = None
        self.auth_status = AuthenticationState.PENDING
        self.error = None

        self.ping_timer = None
        self.missed_pongs = 0


    def ping(self, message):
        message = {
            "event": "ping",
            "data": {
                "state": message
            }
        }
        self.__send(message)

    def subscribe(self, params):

        if self.mode == Mode.Speed and (params.get('channel') == 'aggregates' or params.get('channel') == 'candles'):
            raise Exception("speed mode don't support aggregates and candles channel")

        message = {
            "event": "subscribe",
            "data": params
        }
        self.__send(message)

    def unsubscribe(self, params):
        message = {
            "event": "unsubscribe",
            "data": params
        }
        self.__send(message)

    def subscriptions(self):
        message = {
            "event": "subscriptions"
        }
        self.__send(message)

    def __handle_authentication(self):
        if self.config.get('api_key'):
            auth_info = {
                'event': 'auth',
                'data': {
                    'apikey': self.config['api_key']
                }
            }
        elif self.config.get('bearer_token'):
            auth_info = {
                'event': 'auth',
                'data': {
                    'token': self.config['bearer_token']
                }
            }
        elif self.config.get('sdk_token'):
            auth_info = {
                'event': 'auth',
                'data': {
                    'sdkToken': self.config['sdk_token']
                }
            }
        else:
            self.auth_status = AuthenticationState.UNAUTHENTICATED
            self.error = Exception(MISSING_CREDENTIALS_MESSAGE)

        self.__send(auth_info)
        self.auth_status = AuthenticationState.AUTHENTICATING
        self.auth_timer = Timer(5, self.__check_auth_timeout)
        self.auth_timer.start()

    def __send(self, message):
        self.__ws.send(json.dumps(message))

    def __on_open(self, ws):
        self.ee.emit(CONNECT_EVENT)

    def __on_close(self, ws, close_status_code, close_msg):
        self.ee.emit(DISCONNECT_EVENT, close_status_code, close_msg)

    def __on_message(self, ws, data):
        message = json.loads(data)
        self.ee.emit(MESSAGE_EVENT, data)
        if message['event'] == AUTHENTICATED_EVENT:
            self.ee.emit(AUTHENTICATED_EVENT, message)
            self.auth_status = AuthenticationState.AUTHENTICATED
        elif message['event'] == ERROR_EVENT:
            if message['data'] and message['data']['message'] == UNAUTHENTICATED_MESSAGE:
                self.ee.emit(UNAUTHENTICATED_EVENT, message)
                self.auth_status = AuthenticationState.UNAUTHENTICATED
                self.error = Exception(UNAUTHENTICATED_MESSAGE)
        elif message['event'] == 'pong':
            # Reset missed_pongs counter when receiving a pong
            self.missed_pongs = 0

    def __on_error(self, ws, error):
        self.ee.emit(ERROR_EVENT, error)

    def on(self, event, listener):
        self.ee.on(event, listener)

    def off(self, event, listener):
        self.ee.off(event, listener)

    def __check_auth_timeout(self):
        if self.auth_status == AuthenticationState.AUTHENTICATING:
            self.auth_status = AuthenticationState.UNAUTHENTICATED
            self.error = Exception(AUTHENTICATION_TIMEOUT_MESSAGE)

    def connect(self):
        Thread(target=self.__ws.run_forever).start()
        while True:
            if self.auth_status in [AuthenticationState.AUTHENTICATED, AuthenticationState.UNAUTHENTICATED]:
                break
        if self.error is not None:
            self.__ws.close()
            self.auth_timer.cancel()
            raise self.error
        
        self.missed_pongs = 0
        self.start_ping()
    
    def check_missed_pongs(self):
        if self.missed_pongs > 2:
            self.disconnect()
            raise Exception("Did not receive pong for 2 consecutive times. Disconnecting...")

    def send_ping(self):
        self.ping("")
        self.missed_pongs += 1
        self.check_missed_pongs()
        self.ping_timer = Timer(30, self.send_ping)
        self.ping_timer.start()


    def start_ping(self):
        self.send_ping()


    def disconnect(self):
        if self.__ws is not None:
            self.__ws.close()
            self.error = None

        if self.auth_timer:
            self.auth_timer.cancel()
            self.auth_timer = None
            self.auth_status = AuthenticationState.PENDING

        if self.ping_timer:
            self.ping_timer.cancel()
            self.ping_timer = None

