import discord
from discord.ext import commands
from discord import Embed
import openpyxl
import os

def run_discord_bot():

    intents = discord.Intents.all()  

    bot = commands.Bot(command_prefix='!', intents=intents)

    token = 'MTE3NDE0NjgxMjgxNTM2NDEwNg.GW696D.hOcf39cYqEHxxx5p8s4gpEnEHiSXkUe4zDneKY'
    important_information_channel = 1174466121659859055

    @bot.event
    async def on_ready():
        print(f'{bot.user} is now running!')

    @bot.command(name='positions_exec_board')
    async def positions_exec_board(ctx):
        workbook = openpyxl.load_workbook('member_list.xlsx')
        sheet = workbook.active
        
        channel = bot.get_channel(important_information_channel)
        
        branch = f"**EXECUTIVE BOARD**\n-----------------------------------------------------------"
        target_sheet_name = "executive_board"

        if target_sheet_name in workbook.sheetnames:
            sheet = workbook[target_sheet_name]
            await channel.send(branch)

            for row in sheet.iter_rows(min_row=2, values_only=True):
                formatted_text = f"**{row[0]}:** {row[1]}"
                await channel.send(formatted_text)

    @bot.command(name='positions_exec_cabinet')
    async def positions_exec_cabinet(ctx):
        workbook = openpyxl.load_workbook('member_list.xlsx')
        sheet = workbook.active
        
        channel = bot.get_channel(important_information_channel)
        
        branch = f"**EXECUTIVE CABINET**\n-----------------------------------------------------------"
        target_sheet_name = "executive_cabinet"

        if target_sheet_name in workbook.sheetnames:
            sheet = workbook[target_sheet_name]
            await channel.send(branch)

            for row in sheet.iter_rows(min_row=2, values_only=True):
                formatted_text = f"**{row[0]}:** {row[1]}"
                await channel.send(formatted_text)

    @bot.command(name='positions_legislative')
    async def positions_legislative(ctx):
        workbook = openpyxl.load_workbook('member_list.xlsx')
        sheet = workbook.active

        channel = bot.get_channel(important_information_channel)

        branch = f"**LEGISLATIVE**\n-----------------------------------------------------------"
        target_sheet_name = "legislative"

        if target_sheet_name in workbook.sheetnames:
            sheet = workbook[target_sheet_name]
            await channel.send(branch)

            for row in sheet.iter_rows(min_row=2, values_only=True):
                formatted_text = f"**{row[0]}:** {row[1]}"
                await channel.send(formatted_text)

    @bot.command(name='positions_judicial')
    async def positions_judicial(ctx):
        workbook = openpyxl.load_workbook('member_list.xlsx')
        sheet = workbook.active
        
        channel = bot.get_channel(important_information_channel)
        
        branch = f"**JUDICIAL**\n-----------------------------------------------------------"
        target_sheet_name = "judicial"

        if target_sheet_name in workbook.sheetnames:
            sheet = workbook[target_sheet_name]
            await channel.send(branch)

            for row in sheet.iter_rows(min_row=2, values_only=True):
                formatted_text = f"**{row[0]}:** {row[1]}"
                await channel.send(formatted_text)

    bot.run(token)
