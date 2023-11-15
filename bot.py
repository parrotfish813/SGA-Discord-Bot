import discord
import responces
import openpyxl
import os

def run_discord_bot():
    bot = discord.Client()

    @bot.event
    async def on_ready():
        print(f'{bot.user} is now running!')

    bot.run(os.getenv("TOKEN"))

    @bot.command(name='positions')
    async def positions(ctx):
        workbook = openpyxl.load_workbook('member_list.xlsx')
        sheet = workbook.active

        for row in sheet.iter_rows(min_row=2, values_only=True):
            formatted_text = f"Name: {row[0]}, Position: {row[1]}, Year: {row[2]}"  
            channel_id = 1174085794231767050  
            channel = bot.get_channel(channel_id)
            await channel.send(formatted_text)

