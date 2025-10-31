from fubon_neo._fubon_neo import FugleRealtime

FUGLE_MARKETDATA_API_REST_BASE_URL = FugleRealtime.realtime_rest_url()
FUGLE_MARKETDATA_API_VERSION = 'v1.0'
FUGLE_MARKETDATA_API_NORMAL_WEBSOCKET_BASE_URL = FugleRealtime.realtime_ws_normal_url()
FUGLE_MARKETDATA_API_NORMAL_VERSION = 'v1.0'
FUGLE_MARKETDATA_API_SPEED_WEBSOCKET_BASE_URL = FugleRealtime.realtime_ws_speed_url()
FUGLE_MARKETDATA_API_SPEED_VERSION = 'v1.0'

CONNECT_EVENT = 'connect'
DISCONNECT_EVENT = 'disconnect'
MESSAGE_EVENT = 'message'
ERROR_EVENT = 'error'
AUTHENTICATED_EVENT = 'authenticated'
UNAUTHENTICATED_EVENT = 'unauthenticated'
UNAUTHENTICATED_MESSAGE = 'Invalid authentication credentials'
AUTHENTICATION_TIMEOUT_MESSAGE = 'authentication timeout'
MISSING_CREDENTIALS_MESSAGE= 'missing authentication credentials'
