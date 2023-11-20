import discord
from discord.ext import commands
from discord import Embed
import os
from dotenv import load_dotenv
import asyncio

load_dotenv('.env')

intents = discord.Intents.all() 
intents.members = True
bot = commands.Bot(command_prefix='!', intents=intents)
token: str = os.getenv("DISCORD_TOKEN")

@bot.event
async def on_ready():

    await bot.change_presence(activity=discord.Activity(type=discord.ActivityType.watching, name="!help"))

    print(f'{bot.user} is now running!') 

async def load():
    for filename in os.listdir('./cogs'):
        if filename.endswith('.py'):
            bot.load_extension(f'cogs.{filename[:-3]}')

async def main():
    await load()
    await bot.start(token)

asyncio.run(main())