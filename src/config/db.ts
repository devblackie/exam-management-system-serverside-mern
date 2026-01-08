// serverside/src/config/db.ts
import mongoose from 'mongoose';
import config from './config';


const connectDB = async () => {
  try {
await mongoose.connect(config.databaseURI);
     //await mongoose.connect(process.env.MONGODB_URI || 'mongodb://localhost:27017/medihub');
    console.log('✅ MongoDB connected');

      // GIVE MONGOOSE TIME TO INITIALIZE MODELS
    await new Promise(resolve => setTimeout(resolve, 1000));
    
  } catch (error) {
    console.error('❌ MongoDB connection error:', error);
    process.exit(1);
  }
};

export default connectDB;