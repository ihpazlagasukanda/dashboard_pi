const client = require('./redisClient');

const CacheUtils = {
    // Generate consistent cache key
    generateKey: (prefix, params) => {
        const keyParts = [prefix];
        const orderedParams = {};
        
        // Urutkan parameter secara konsisten
        ['provinsi', 'kabupaten', 'kecamatan', 'tahun', 'bulan', 'produk', 'status'].forEach(key => {
            if (params[key] !== undefined && params[key] !== null && params[key] !== '') {
                orderedParams[key] = params[key];
            }
        });
        
        for (const [key, value] of Object.entries(orderedParams)) {
            keyParts.push(`${key}:${value}`);
        }
        return keyParts.join(':');
    },

    // Get data from cache
    get: async (key) => {
        try {
            const data = await client.get(key);
            return data ? JSON.parse(data) : null;
        } catch (error) {
            console.log('Cache get error:', error.message);
            return null;
        }
    },

    // Set data to cache
    set: async (key, data, ttl = 300) => {
        try {
            await client.setEx(key, ttl, JSON.stringify(data));
            return true;
        } catch (error) {
            console.log('Cache set error:', error.message);
            return false;
        }
    },

    // Delete cache
    delete: async (key) => {
        try {
            await client.del(key);
            return true;
        } catch (error) {
            console.log('Cache delete error:', error.message);
            return false;
        }
    },

    // Clear cache by pattern
    clearByPattern: async (pattern) => {
        try {
            const keys = await client.keys(pattern);
            if (keys.length > 0) {
                await client.del(keys);
            }
            return keys.length;
        } catch (error) {
            console.log('Cache clear error:', error.message);
            return 0;
        }
    }
};

module.exports = CacheUtils;