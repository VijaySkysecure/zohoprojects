const mongoose = require("mongoose");
const chalk = require("cli-color");
const { mongoDBConnectionString } = require("./config")

/**
 * @description This class handles the mongoDb connection related works
 */
class MongoDbConnection {
    constructor() {
        mongoose.connect(mongoDBConnectionString);
        mongoose.connection.once('open', () => console.log(chalk.greenBright.bold.italic('MongoDb connection opened...')));
        mongoose.connection.on('connected', () => console.log(chalk.cyanBright.bold.italic('MongoDb connected successfully...')));
        mongoose.connection.on('reconnected', () => console.log(chalk.blueBright.bold.italic('MongoDb reconnected successfully...')));
        mongoose.connection.on('disconnected', () => console.log(chalk.redBright.bold.italic('MongoDb disconnected!!!')));
        mongoose.connection.on('error', () => console.log(chalk.redBright.bold.italic('MongoDb connection error!!!')));
    }
}

module.exports = MongoDbConnection