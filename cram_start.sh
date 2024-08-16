nohup node server.js > nohup.out 2>&1 &
echo $! > server1.pid
echo "Server started with PID $(cat server1.pid)"
