import discord
from discord.ext import commands
from discord import Embed
import openpyxl
import os
from dotenv import load_dotenv

load_dotenv('.env')

def run_discord_bot():

    intents = discord.Intents.all()  

    bot = commands.Bot(command_prefix='!', intents=intents)

    token: str = os.getenv("DISCORD_TOKEN")
    # Convert the channel IDs to integers
    important_information_channel = int(os.getenv("IMPORTANT_INFORMATION_CHANNEL"))
    exec_information_channel = int(os.getenv("EXEC_INFORMATION_CHANNEL"))
    legislative_information_channel = int(os.getenv("LEGISLATIVE_INFORMATION_CHANNEL"))  
    judicial_information_channel = int(os.getenv("JUDICIAL_INFORMATION_CHANNEL"))
    lec_information_channel  = int(os.getenv("LEC_INFORMATION_CHANNEL"))
    abc_information_channel = int(os.getenv("ABC_INFORMATION_CHANNEL"))
    acc_information_channel = int(os.getenv("ACC_INFORMATION_CHANNEL"))
    pec_information_channel = int(os.getenv("PEC_INFORMATION_CHANNEL"))
    media_information_channel = int(os.getenv("MEDIA_INFORMATION_CHANNEL"))

    print(f"important_information_channel: {important_information_channel}")

    @bot.event
    async def on_ready():
        print(f'{bot.user} is now running!')

    """ POSITIONS COMMANDS """

    @bot.command(name='positions_exec_board')
    async def positions_exec_board(ctx):
        workbook = openpyxl.load_workbook('member_list.xlsx')
        sheet = workbook.active
        
        channel = bot.get_channel(important_information_channel)
        
        branch = f"-----------------------------------------------------------\n**EXECUTIVE BOARD**\n-----------------------------------------------------------"
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
        
        branch = f"-----------------------------------------------------------\n**EXECUTIVE CABINET**\n-----------------------------------------------------------"
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

        branch = f"-----------------------------------------------------------\n**LEGISLATIVE**\n-----------------------------------------------------------"
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
        
        branch = f"-----------------------------------------------------------\n**JUDICIAL**\n-----------------------------------------------------------"
        target_sheet_name = "judicial"

        if target_sheet_name in workbook.sheetnames:
            sheet = workbook[target_sheet_name]
            await channel.send(branch)

            for row in sheet.iter_rows(min_row=2, values_only=True):
                formatted_text = f"**{row[0]}:** {row[1]}"
                await channel.send(formatted_text)

    
    """ INFORMATIONAL COMMANDS"""

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

    @bot.command(name='branch_info_lec')
    async def branch_info_lec(ctx):
        channel = bot.get_channel(lec_information_channel)

        formatted_text = f"The Florida Polytechnic University SGA Legislative Executive Committee (LEC) is comprised of the SGA Senate President, SGA Senate Pro Tempore, and all Committee Chairs.\n\nLEC is where all senate chairs discuss important updates from their respective committees, progress on individual committee projects, as well as any issues that may have come up during their respective meetings that needs to be addressed.\n\nPotential issues that students are facing are brought to this committee, and LEC determines what direction the individual committees can take to address these concerns, either by finding resources or by determining possible solutions."
        await channel.send(formatted_text)

    @bot.command(name='branch_info_abc')
    async def branch_info_abc(ctx):
        channel = bot.get_channel(abc_information_channel)

        formatted_text = f""
        await channel.send(formatted_text)

    @bot.command(name='branch_info_acc')
    async def branch_info_acc(ctx):
        channel = bot.get_channel(acc_information_channel)

        formatted_text = f""
        await channel.send(formatted_text)

    @bot.command(name='branch_info_pec')
    async def branch_info_pec(ctx):
        channel = bot.get_channel(pec_information_channel)

        formatted_text = f""
        await channel.send(formatted_text)

    @bot.command(name='branch_info_media')
    async def branch_info_media(ctx):
        channel = bot.get_channel(media_information_channel)

        formatted_text = f"The Florida Polytechnic University SGA Department of Student Media is comprised of the SGA Director of Student Media and additional deputies, such as the Deputy of Communication, the Deputy of Marketing, and the Deputy of Enforcement.\n\nThe Department of Student Media is responsible for ensuring that the student body is aware of SGA and SGA’s outreach to students, along with supporting media groups, on campus.\n\nThis department oversees any and all social media, as well as create flyers and merchandise, for SGA and RSOs."
        await channel.send(formatted_text)

    bot.run(token)
