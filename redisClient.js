const redis = require('redis');

class RedisClient {
    constructor() {
        this.client = null;
        this.isConnected = false;
        this.connect();
    }

    connect() {
        try {
            // Konfigurasi berdasarkan environment
            const isProduction = process.env.NODE_ENV === 'production';
            
            const config = {
                socket: {
                    host: process.env.REDIS_HOST || 'localhost',
                    port: process.env.REDIS_PORT || 6379,
                    reconnectStrategy: (retries) => {
                        if (retries > 10) {
                            console.log('Too many retries on Redis. Connection terminated');
                            return new Error('Too many retries');
                        }
                        return Math.min(retries * 100, 3000);
                    }
                },
                password: process.env.REDIS_PASSWORD || null,
                database: parseInt(process.env.REDIS_DB) || 0
            };

            this.client = redis.createClient(config);

            this.client.on('connect', () => {
                console.log('Redis client connected');
                this.isConnected = true;
            });

            this.client.on('error', (err) => {
                console.error('Redis client error:', err.message);
                this.isConnected = false;
            });

            this.client.on('end', () => {
                console.log('Redis client disconnected');
                this.isConnected = false;
            });

            this.client.connect();

        } catch (error) {
            console.error('Failed to create Redis client:', error);
            this.isConnected = false;
        }
    }

    async get(key) {
        try {
            if (!this.client || !this.isConnected) {
                return null;
            }
            return await this.client.get(key);
        } catch (error) {
            console.error('Redis get error:', error);
            return null;
        }
    }

    async setEx(key, ttl, value) {
        try {
            if (!this.client || !this.isConnected) {
                return false;
            }
            await this.client.setEx(key, ttl, value);
            return true;
        } catch (error) {
            console.error('Redis setEx error:', error);
            return false;
        }
    }

    async del(key) {
        try {
            if (!this.client || !this.isConnected) {
                return false;
            }
            await this.client.del(key);
            return true;
        } catch (error) {
            console.error('Redis del error:', error);
            return false;
        }
    }

    async keys(pattern) {
        try {
            if (!this.client || !this.isConnected) {
                return [];
            }
            return await this.client.keys(pattern);
        } catch (error) {
            console.error('Redis keys error:', error);
            return [];
        }
    }

    async quit() {
        try {
            if (this.client && this.isConnected) {
                await this.client.quit();
                this.isConnected = false;
            }
        } catch (error) {
            console.error('Redis quit error:', error);
        }
    }
}

// Singleton instance
const redisClient = new RedisClient();

// Graceful shutdown
process.on('SIGINT', async () => {
    await redisClient.quit();
    process.exit(0);
});

process.on('SIGTERM', async () => {
    await redisClient.quit();
    process.exit(0);
});

module.exports = redisClient;