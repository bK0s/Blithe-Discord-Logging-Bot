import discord
from discord.ext import commands, tasks
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
#from matplotlib import pyplot as plt
import datetime
import pytz
import os

#IDs and Constants
SPREADSHEET_ID = '<spreadsheet_id>' #String
RANGE = '<range>' #String 
CHANNEL_ID = '<channel_id>' #int
channel = None

creds = Credentials.from_authorized_user_file('token.json', ["https://www.googleapis.com/auth/spreadsheets"])
sheets = build('sheets', 'v4', credentials=creds).spreadsheets()

# Discord Bot Config Setup 
intents = discord.Intents(messages = True, guilds = True, reactions = True, members = False, presences = False)
bot = commands.Bot(command_prefix = os.environ['PREFIX'], intents = intents)

@bot.event
async def on_ready():
    print('Bot is online')
    global channel
    channel = await bot.fetch_channel(CHANNEL_ID)
    credential_refresher.start()
    print('Ready')

# Google Sheets API setup/initialization. Credentials will automatically refresh every hour
@tasks.loop(seconds=59, minutes=59)
async def credential_refresher():
    global creds
    global sheets
    creds.refresh(Request())
    if not creds.valid:
        print('Creds not valid')
    with open('token.json', 'w') as token:
        token.write(creds.to_json())
    sheets = build('sheets', 'v4', credentials=creds).spreadsheets()
    print("Refreshed Google API credentials")



# Handles the core logging whenever ticket bot posts a transcript
@bot.event
async def on_message_edit(_, message): #Ticket bot always edits the message after sending, so this only runs after the edited, finalized data is in
    # Ensure message is in transcripts channel and sent by ticket bot
    if message.channel.id != CHANNEL_ID or message.author.id != "<BOT_ID>": 
        return

    embed = message.embeds[0]
    entry = [[]]

    #Make sure ticket number hasn't already been logged
    ticket_num = embed.fields[1].value[embed.fields[1].value.index('-')+1:]
    #TODO: Fix Sheets Key error
    logged_nums = sheets.values().get(spreadsheetId=SPREADSHEET_ID, range="C2:C10000",
                            majorDimension = 'COLUMNS').execute()['values'][0]
    if ticket_num in logged_nums: return

    #Datetime of ticket close (Convert to Mountain timezone)
    entry[0].append(message.created_at.replace(tzinfo=datetime.timezone.utc).astimezone(pytz.timezone('Canada/Mountain')).strftime(r'%m/%d/%Y %H:%M:%S'))
    
    #Discord ID (strip tag notation, e.g. <@123456789> --> 123456789)
    entry[0].append(embed.fields[0].value[2:-1])
    
    #Ticket number (strip status and add ' to prevent formatting e.g. closed-0670 -> '0670)
    entry[0].append("'" + ticket_num)

    #Skip approval status (this is handled later to allow for delayed manual reactions)
    entry[0].append("")

    #Transcript Link (strip hyperlink notation)
    entry[0].append(embed.fields[3].value[embed.fields[3].value.index('(')+1:-1])

    #Log values in the spreadsheet
    sheets.values().append(spreadsheetId=SPREADSHEET_ID, valueInputOption='USER_ENTERED',
                            range=RANGE, body={'values': entry}).execute()
    
    #Add reactions options for approval status
    await message.add_reaction('✅')
    await message.add_reaction('❌')
    await message.add_reaction('❕')


    print(f'Logged Ticket #{ticket_num}. Now awaiting approval reaction...')
    


# Handles the 'Approval Status' section based on staff reactions to transcript. Re-reacting will update the value in the spreadsheet
@bot.event
async def on_raw_reaction_add(reaction_event):
    # Ensure reaction is NOT from this bot and that the message is in transcripts channel and sent by Ticket Bot
    if reaction_event.user_id == "<TICKET_BOT_ID>" or reaction_event.channel_id != CHANNEL_ID:
        return
    message = await channel.fetch_message(reaction_event.message_id)
    if message.author.id != '<BOT_ID>':
        return

    # Grab approval status from reaction
    referral = ""
    approval = ""
    if str(reaction_event.emoji) == "✅": 
        approval = "accepted"
        await message.add_reaction('🟣') #REDDIT
        await message.add_reaction('🔵') #REFERRAL
        await message.add_reaction('🟡') #PMC
        # await channel.send()
    elif str(reaction_event.emoji) == "❌":
        approval = "denied"
    elif str(reaction_event.emoji) == "❕":
        approval = 'incomplete'
    elif str(reaction_event.emoji) == '🟣':
        referral = 'option1'
    elif str(reaction_event.emoji) == '🔵':
        referral = 'option2'
    elif str(reaction_event.emoji) == '🟡':
        referral = 'option3'
    else:
        print('Invalid reaction detected')
        return

    # Grab ticket number from message
    ticket_num = message.embeds[0].fields[1].value[message.embeds[0].fields[1].value.index('-')+1:]

    # Locate row from ticket number
    row = sheets.values().get(spreadsheetId=SPREADSHEET_ID, range="C2:C10000",
                                    majorDimension = 'COLUMNS').execute()['values'][0].index(ticket_num)+2
    
    # Add new Approval Status value to corresponding cell
    if approval != '':
        sheets.values().update(spreadsheetId=SPREADSHEET_ID, range=f'D{row}', valueInputOption='USER_ENTERED', body={'values':[[approval]]}).execute()
        print(f'Set approval for to Ticket #{ticket_num} to {approval}')
    # Add Referral Status value to corresponding cell
    if referral != '':
        sheets.values().update(spreadsheetId=SPREADSHEET_ID, range=f'F{row}', valueInputOption='USER_ENTERED', body={'values':[[referral]]}).execute()
        print(f'Set Ticket #{ticket_num} referral to {referral}')

#The search function to retrieve entries. 
@bot.command(aliases=['Lookup', 'search'])
async def lookup(ctx, user: discord.Member):
    message = f"Here are all of the stored tickets opened by {user.mention}: "
    #Search through database for entries with given user's id
    response = sheets.values().get(spreadsheetId=SPREADSHEET_ID, range="A2:F10000").execute()
    values = response.get('values', [])
    if not values:
        print('No Data found')
        return
    #There may be mutliple, so construct/send each individually then keep searching
    for row in values:
        if row[1] == str(user.id):
            embed = discord.Embed()
            embed.set_author(name = f'{user.name}#{user.discriminator}', icon_url=user.avatar_url)
            embed.add_field(name = "Ticket Owner", value = user.mention)
            embed.add_field(name = "Ticket Number", value = row[2])
            embed.add_field(name = "Date/time(MT)", value = row[0])
            embed.add_field(name = 'Direct Transcript', value = f'[Direct Transcript]({row[4]})')
            if row[3] == 'accepted':
                embed.color = discord.Color.green()
                embed.add_field(name = 'Approval Status', value = 'Accepted ✅')
            elif row[3] == 'denied':
                embed.color = discord.Color.red()
                embed.add_field(name = 'Approval Status', value = 'Denied ❌')
            else:
                embed.color = discord.Color.dark_gray()
                embed.add_field(name = 'Approval Status', value = 'Pending')

            await ctx.send(message, embed=embed)
            message = ""
    
    if message:
        await ctx.send(f"Could not find any tickets opened by {user.mention}")
    print(f"Completed Lookup on {user.mention}")

#The Stats function to display graph of user referals
@bot.command(aliases = ['Stats', 'analytics'])
async def stats(ctx, message):
    if ctx.channel.id != "<STAFF_CHANNEL>":
        await ctx.send('Command not available from this channel')
        return
    print(message)
    response = sheets.values().get(spreadsheetId=SPREADSHEET_ID, range="A2:F10000").execute()
    values = response.get('values', [])
    embed = discord.Embed()
    if message == 'referral':
        await ctx.send('Gathering referral data')
        embed.color = discord.Color.blue()
        option1= 0
        option2= 0
        option3= 0
        for row in values:
            if row[5] == 'option1':
                option1+=1
            elif row[5] == 'option2':
                option2 +=1
            elif row[5]== 'option3':
                option3 +=1
            else:
                pass
        embed.add_field(name='REFERRAL DATA', value=f'Number of players from Option1: {option1}\nNumber of players from Option2: {option2}\nNumber of players from Option3: {option3}\n\n')
        embed.add_field(name='NOTE', value=f'This data was generated using {option1+option2+option3} applications')
    else:
        await ctx.send('Incorrect command. Use: /stats referral')
        return

    await ctx.send(embed=embed)
#TODO: Implement Graph Embed

#     graph = plt.figure(figsize = (10,7))   
#     titles = []
#     data = '' #get data from spreadsheats
#     plt.pie(data, lables = titles)
#     path_to_file = ""
#     fname = ''
#     file = discord.File(f"{path_to_file}", filename=f'{fname}')
#     embed = discord.Embed(title='Analytics', color=discord.Colour.blue())
#     await ctx.send(embed=embed, file=file)
    

bot.run(os.environ['TOKEN'])