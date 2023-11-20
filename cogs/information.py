

import discord
from discord.ext import commands
from discord import Embed
import openpyxl
import os
from dotenv import load_dotenv
import functions.doc_format as doc_format

load_dotenv('.env')

# EXECUTIVE BRANCH CHANNELS
exec_branch_information_channel = int(os.getenv("EXEC_BRANCH_CHANNEL"))
nova_information_channel = int(os.getenv("NOVA_INFORMATION_CHANNEL"))
internal_affairs_information_channel = int(os.getenv("INTERNAL_AFFAIRS_INFORMATION_CHANNEL"))
external_affairs_information_channel = int(os.getenv("EXTERNAL_AFFAIRS_INFORMATION_CHANNEL"))
student_media_information_channel = int(os.getenv("STUDENT_MEDIA_INFORMATION_CHANNEL"))

# LEGISLATIVE BRANCH CHANNELS
legislative_branch_information_channel = int(os.getenv("LEGISLATIVE_BRANCH_CHANNEL"))  
lec_information_channel  = int(os.getenv("LEC_INFORMATION_CHANNEL"))
abc_information_channel = int(os.getenv("ABC_INFORMATION_CHANNEL"))
acc_information_channel = int(os.getenv("ACC_INFORMATION_CHANNEL"))
pec_information_channel = int(os.getenv("PEC_INFORMATION_CHANNEL"))

# JUDICIAL BRANCH CHANNELS
judicial_branch_information_channel = int(os.getenv("JUDICIAL_BRANCH_CHANNEL"))
judicial_information_channel = int(os.getenv("JUDICIAL_INFORMATION_CHANNEL"))

# EXECUTIVE BRANCH DOCUMENTS
exec_information_document = doc_format.get_data_from_word_double("files/exec/exec_info.docx")
nova_information_document = doc_format.get_data_from_word_double("files/exec/nova_info.docx")
internal_affairs_information_document = doc_format.get_data_from_word_double("files/exec/internal_affairs_info.docx")
external_affairs_information_document = doc_format.get_data_from_word_double("files/exec/external_affairs_info.docx")
student_media_information_document = doc_format.get_data_from_word_double("files/exec/student_media_info.docx")

# LEGISLATIVE BRANCH DOCUMENTS
legislative_branch_information_document = doc_format.get_data_from_word_double("files/legislative/legislative_info.docx")
lec_information_document = doc_format.get_data_from_word_double("files/legislative/lec_info.docx")
abc_information_document = doc_format.get_data_from_word_double("files/legislative/abc_info.docx")
acc_information_document = doc_format.get_data_from_word_double("files/legislative/acc_info.docx")
pec_information_document = doc_format.get_data_from_word_double("files/legislative/pec_info.docx")

# JUDICIAL BRANCH DOCUMENTS
judicial_branch_information_document = doc_format.get_data_from_word_double("files/judicial/judicial_branch_info.docx")
judicial_information_document = doc_format.get_data_from_word_double("files/judicial/judicial_info.docx")

class information(commands.Cog):
    def __init__(self, bot):
        self.bot = bot

        async def on_ready():
            print(f'{bot.user} is now running!')

        """ INFORMATIONAL COMMANDS"""

        # EXECUTIVE BRANCH COMMANDS
        @commands.command(name='branch_info_exec')
        async def branch_info_exec(self, ctx):
            channel = bot.get_channel(exec_branch_information_channel)

            await channel.send(exec_information_document)

        @commands.command(name='nova_info')
        async def nova_info(self, ctx):
            channel = bot.get_channel(1175909206541488208)

            await channel.send(nova_information_document)

        @commands.command(name='internal_affairs_info')
        async def internal_affairs_info(self, ctx):
            channel = bot.get_channel(internal_affairs_information_channel)

            await channel.send(internal_affairs_information_document)

        @commands.command(name='external_affairs_info')
        async def external_affairs_info(self, ctx):
            channel = bot.get_channel(external_affairs_information_channel)

            await channel.send(external_affairs_information_document)

        @commands.command(name='student_media_info')
        async def student_media_info(self, ctx):
            channel = bot.get_channel(student_media_information_channel)

            await channel.send(student_media_information_document)

        # LEGISLATIVE BRANCH COMMANDS

        @commands.command(name='branch_info_legislative')
        async def branch_info_legislative(self, ctx):
            channel = bot.get_channel(legislative_branch_information_channel)

            await channel.send(legislative_branch_information_document)

        @commands.command(name='branch_info_lec')
        async def lec_info(self, ctx):
            channel = bot.get_channel(lec_information_channel)

            await channel.send(lec_information_document)

        @commands.command(name='abc_info')
        async def abc_info(self, ctx):
            channel = bot.get_channel(abc_information_channel)

            await channel.send(abc_information_document)

        @commands.command(name='acc_info')
        async def acc_info(self, ctx):
            channel = bot.get_channel(acc_information_channel)

            await channel.send(acc_information_document)

        @commands.command(name='pec_info')
        async def pec_info(self, ctx):
            channel = bot.get_channel(pec_information_channel)

            await channel.send(pec_information_document)

        # JUDICIAL BRANCH COMMANDS

        @commands.command(name='branch_info_judicial')
        async def branch_info_judicial(self, ctx):
            channel = bot.get_channel(judicial_branch_information_channel)

            await channel.send(judicial_branch_information_document)

        @commands.command(name='judicial_info')
        async def judicial_info(self, ctx):
            channel = bot.get_channel(judicial_information_channel)

            await channel.send(judicial_information_document)

async def setup(bot):
    await bot.add_cog(information(bot))