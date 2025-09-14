const redis = require('redis');

// Buat koneksi Redis
const client = redis.createClient({
  socket: {
    host: 'localhost',
    port: 6379
  },
  password: null // Jika tidak pakai password
});

// Handle error connection
client.on('error', (err) => {
  console.error('Redis Error:', err);
});

// Handle successful connection
client.on('connect', () => {
  console.log('Connected to Redis successfully');
});

// Connect to Redis
(async () => {
  try {
    await client.connect();
    console.log('Redis client connected');
  } catch (error) {
    console.error('Redis connection failed:', error);
  }
})();

module.exports = client;