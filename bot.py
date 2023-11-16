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
    exec_information_channel = 1151911705610289203
    legislative_information_channel = 1151912030828245062
    judicial_information_channel = 1151912235560615999


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

    @bot.command(name='branch_info_exec')
    async def branch_info_exec(ctx):
        channel = bot.get_channel(exec_information_channel)

        formatted_text = f"The Student Government Association's (SGA) Executive Branch is comprised of the SGA Executive Board and the SGA Executive Cabinet.\n\nThe Florida Polytechnic University SGA Executive Board represents the executive authority of the Student Government Association, including in matters relating to the governance and wellbeing of the student body. They represent the student voice to the university administration and is comprised of the Student Body President, Vice President, Treasurer, and Chief of Staff.\n\nThe Florida Polytechnic University SGA Executive Cabinet assists the president in conduct of business dependent on their respective areas. Information about the departments can be found in their individual channels below, and information about the agencies can be found in their separate Discords, which can be found in the other-servers channel.\n\nStudent Body President email: sga-president@floridapoly.edu"
        await channel.send(formatted_text)

    @bot.command(name='branch_info_legislative')
    async def branch_info_legislative(ctx):
        channel = bot.get_channel(legislative_information_channel)

        formatted_text = f"The Student Government Association's (SGA) Legislative Branch is comprised of the SGA General Senate, SGA Legislative Executive Committee (LEC), SGA Policies & Ethics Committee (PEC),  SGA Auditing & Budgeting Committee (ABC), and SGA Advocacy & Communications Committee (ACC).\n\nThe Florida Polytechnic University SGA General Senate is comprised of your elected Senators who draft legislation and create programs to enhance and enrich the experience of all students. Senate is made up of 15 total representatives that each represent different parts of the student body from On-Campus Representative to Freshman Representative to Engineering Representative.\n\nEach representative must be a member of at least one of the committees (PEC, ABC, ACC) as part of their duties as a senator. Those that are chairs of each committee must also take part in LEC. For more information on the different committees, and what they do, please refer to their respective channels below!\n\nLegislative Branch email: sga-senate@floridapoly.edu "
        await channel.send(formatted_text)

    @bot.command(name='branch_info_judicial')
    async def branch_info_legislative(ctx):
        channel = bot.get_channel(judicial_information_channel)

        formatted_text = f"The Student Government Association's (SGA) Judicial Branch is comprised of the SGA Supreme Court and the SGA Office of Student Counsel.\n\nThe Florida Polytechnic University SGA Supreme Court represents the judicial authority of the Student Government Association, including in matters relating to the interpretation of the Constitution and related statutes. There are seven Justices on the Supreme Court, including the Chief Justice, Raking Justice, Senior Justice, and four Associate Justices.\n\nThe Florida Polytechnic University SGA Office of Student Counsel provides the student body access to professional counselors during court and administrative proceedings. The Office of Student Counsel provides assistance to students in matters pertaining to alleged Student Code of Conduct violations.\n\nJudicial Branch email: sga-judicial@floridapoly.edu"
        await channel.send(formatted_text)

    bot.run(token)
