"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
// serverside/src/config/db.ts
const mongoose_1 = __importDefault(require("mongoose"));
const config_1 = __importDefault(require("./config"));
const connectDB = async () => {
    try {
        await mongoose_1.default.connect(config_1.default.databaseURI);
        //await mongoose.connect(process.env.MONGODB_URI || 'mongodb://localhost:27017/medihub');
        console.log('✅ MongoDB connected');
        // GIVE MONGOOSE TIME TO INITIALIZE MODELS
        await new Promise(resolve => setTimeout(resolve, 1000));
    }
    catch (error) {
        console.error('❌ MongoDB connection error:', error);
        process.exit(1);
    }
};
exports.default = connectDB;
