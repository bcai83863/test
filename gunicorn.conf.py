# load app once in master; workers fork (avoids per-worker font cache rebuild)
preload_app = True
timeout = 120        # give sync workers more time for heavy matplotlib operations
workers = 1
bind = "0.0.0.0:10000"
