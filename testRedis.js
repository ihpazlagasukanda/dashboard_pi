const client = require('./redisClient');

async function testRedis() {
  try {
    // Test koneksi
    await client.set('test', 'Redis is working!');
    const value = await client.get('test');
    console.log('Redis test value:', value);
    
    // Test dengan data JSON
    const testData = { message: 'Hello Redis', timestamp: new Date() };
    await client.setEx('test:json', 60, JSON.stringify(testData));
    
    const jsonValue = await client.get('test:json');
    console.log('Redis JSON test:', JSON.parse(jsonValue));
    
  } catch (error) {
    console.error('Redis test failed:', error);
  }
}

testRedis();