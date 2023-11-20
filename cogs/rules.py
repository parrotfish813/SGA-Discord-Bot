import discord
from discord.ext import commands
from discord import Embed
import openpyxl
import os
from dotenv import load_dotenv
from functions import doc_format as doc_format

load_dotenv('.env')

# RULES CHANNEL 
rules_channel = int(os.getenv("RULES_CHANNEL"))

# RULES
rules_document = doc_format.get_data_from_word_single("files/rules.docx")

class rules(commands.Cog):
    def __init__(self, bot):
        self.bot = bot

        async def on_ready():
            print(f'{bot.user} is now running!')

        # RULES COMMANDS
        @commands.command(name='rules')
        async def rules(self, ctx):
            channel = bot.get_channel(1175909206541488208)

            await channel.send(rules_document)

async def setup(bot):
    await bot.add_cog(rules(bot))