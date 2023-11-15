import discord
from discord.ext import commands
import openpyxl
import os

def run_discord_bot():
    intents = discord.Intents.all()  

    bot = commands.Bot(command_prefix='!', intents=intents)

    # bot = discord.Client()

    @bot.event
    async def on_ready():
        print(f'{bot.user} is now running!')

    @bot.command(name='positions')
    async def positions(ctx):
        workbook = openpyxl.load_workbook('member_list.xlsx')
        sheet = workbook.active

        for row in sheet.iter_rows(min_row=2, values_only=True):
            formatted_text = f"Name: {row[0]}, Position: {row[1]}, Year: {row[2]}"  
            channel = bot.get_channel(os.getenv("IMPORT_INFORMATION_CHANNEL_ID"))
            await channel.send(formatted_text)

    bot.run(os.getenv("TOKEN"))

