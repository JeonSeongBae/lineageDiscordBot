#system
import sys
import os
import win32con
import win32api
import win32gui
import asyncio
import datetime
import random
import re
import time
import logging

#discord
import discord
from discord.ext import commands
from discord.ext.commands import CommandNotFound
from gtts import gTTS

#google
from oauth2client.service_account import ServiceAccountCredentials
import gspread

if not discord.opus.is_loaded():
	discord.opus.load_opus('opus')

client = discord.Client()
client = commands.Bot(command_prefix="", help_command = None, description='일상디코봇')

##
basicSetting = []
##


def init():
	global basicSetting
	global command
	global credentials

	command = []

	command_inidata = open('command.ini', 'r', encoding = 'utf-8')
	tmp_command_inidata = command_inidata.readlines()
	command_inputData = tmp_command_inidata

	inidata = open('config.ini', 'r', encoding = 'utf-8')
	tmp_inputData = inidata.readlines()
	inputData = tmp_inputData
	basicSetting.append(inputData[0][0:])     #basicSetting[0] : bot_token

	for i in range(len(command_inputData)):
		tmp_command = command_inputData[i][0:].rstrip('\n')
		fc = tmp_command.split(', ')
		command.append(fc)
		fc = []

	scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
	credentials = ServiceAccountCredentials.from_json_keyfile_name(
			'./chaos-280903-9cc7510906e6.json', scope)

init()
token = basicSetting[0]

def handle_exit():
	client.loop.run_until_complete(client.logout())

	for t in asyncio.Task.all_tasks(loop=client.loop):
		if t.done():
			try:
				t.exception()
			except asyncio.CancelledError:
				continue
			continue
		t.cancel()
		try:
			client.loop.run_until_complete(asyncio.wait_for(t, 5, loop=client.loop))
			t.exception()
		except asyncio.InvalidStateError:
			pass
		except asyncio.TimeoutError:
			pass
		except asyncio.CancelledError:
			pass

def connectDiscord():
	try:
		client.loop.run_until_complete(client.start(token))
	except SystemExit:
		handle_exit()
	except KeyboardInterrupt:
		handle_exit()

async def task():
	await client.wait_until_ready()
	await asyncio.sleep(1) # task runs every 60 seconds

while True:
	#로그인
	@client.command(name=command[0][0], aliases=command[0][1:])
	async def loginBot(ctx):
		return

	#채팅방
	@client.command(name=command[1][0], aliases=command[1][1:])
	async def connectVoiceChatting(ctx):
		if ctx.voice_client is None:
			if ctx.author.voice:
				await ctx.author.voice.channel.connect(reconnect = True)
			else:
				await ctx.send('Please Enter Voice Channel', tts=False)
				return
		else:
			if ctx.voice_client.is_playing():
				ctx.voice_client.stop()
			await ctx.voice_client.move_to(ctx.author.voice.channel)
		
		await ctx.send('Enter Bot in ' + client.get_channel(ctx.author.voice.channel.id).name, tts=False)
		return

	#시간표
	@client.command(name=command[2][0], aliases=command[2][1:])
	async def printSchedule(ctx):
		gc1 = gspread.authorize(credentials).open("보스타임").worksheet('2.보스 리젠 시간').range('AB3:AF29')

		schedule = []
		outPut = []

		for cell in gc1:
			schedule[cell.row].append(cell.value + '\t')
		
		for i in len(schedule):
			for j in len(schedule[i]):
				outPut.append(schedule[i][j])
			outPut.append('\n')

		embed = discord.Embed(
			title = '',
			description= outPut,
			color=0x0000ff
		)
		await ctx.send( embed=embed, tts=False)

		return

	@client.event
	async def on_command_error(ctx, error):
		if isinstance(error, CommandNotFound):
			return
		elif isinstance(error, discord.ext.commands.MissingRequiredArgument):
			return
		raise error

	client.loop.create_task(task())

	print("connectDiscord")
	connectDiscord()
