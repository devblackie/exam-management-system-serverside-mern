import crypto from 'crypto';

const secret = crypto.randomBytes(32).toString('hex');
console.log('Generated JWT_SECRET:', secret);
//how to use this secret in your application
// node generateSecret.js