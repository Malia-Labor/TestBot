using Discord;
using Discord.Commands;
using Google.Apis.Sheets.v4.Data;
using BusinessLogic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BasicBot.Modules
{
    public class Signup : ModuleBase<SocketCommandContext>
    {
        private SignupHandler signupHandler = new SignupHandler();

        // Adds user's name to google sheets, overwrites data for user's row if user already exists in table
        [Command("signup")]
        [Alias("su")]
        [Summary("Sign up for an event")]
        public async Task AddToSignup([Summary("Name of the sheet to add to.")] string sheetName, [Remainder, Summary("Extra information to add with name.")] string message)
        {
            if (sheetName.ToLower() == "template")
            {
                await ReplyAsync("Cannot access template.");
                return;
            }
            bool success = await signupHandler.AddToSignup(sheetName, Context.User.Username, Context.User.Id.ToString(), message); ;
            if (success)
                await Context.Channel.SendMessageAsync("Success");
            else
                await Context.Channel.SendMessageAsync("Failed");
        }

        // Removes user from signup, at the moment success returns false only if an error occurred in the delete, will be true if user is not in sheet
        [Command("withdraw")]
        [Alias("w", "remove")]
        [Summary("Withdraw from an event.")]
        public async Task RemoveFromSignup([Summary("Name of the sheet to delete from.")] string sheetName)
        {
            if (sheetName.ToLower() == "template")
            {
                await ReplyAsync("Cannot access template.");
                return;
            }
            bool success = await signupHandler.RemoveFromSignup(sheetName, Context.User.Id.ToString());
            if (success)
                await Context.Channel.SendMessageAsync("Success");
            else
                await Context.Channel.SendMessageAsync("Failed");
        }

        [Command("create")]
        [Alias("c")]
        [RequireUserPermission(GuildPermission.Administrator)]
        [Summary("Create a sheet")]
        public async Task CreateSignup([Remainder, Summary("Name for new sheet.")] string sheetName)
        {
            bool success = await signupHandler.CopyTemplateAsync(sheetName);
            if (success)
                await Context.Channel.SendMessageAsync("Success");
            else
                await Context.Channel.SendMessageAsync("Failed");
        }

        [Command("delete")]
        [Alias("d")]
        [RequireUserPermission(GuildPermission.Administrator)]
        [Summary("Deletes a sheet")]
        public async Task ClearSignup([Remainder, Summary("Name for sheet to delete.")] string sheetName)
        {
            if (sheetName.ToLower() != "template")
            {
                bool success = await signupHandler.DeleteSheetAsync(sheetName);
                if (success)
                    await Context.Channel.SendMessageAsync("Success");
                else
                    await Context.Channel.SendMessageAsync("Failed");
            }
            else
                await Context.Channel.SendMessageAsync("Cannot delete template.");
        }

        [Command("read")]
        [Summary("Read data from google sheets.")]
        public async Task ReadSheets([Remainder, Summary("Name of sheet to read.")] string sheetName)
        {
            var values = await signupHandler.ReadAsync($"{sheetName}!A:C");
            string reply = "";
            foreach (var row in values.Skip(1))
            {
                if (row.FirstOrDefault() != null)
                {
                    reply += string.Join(" ", row.Select(r => r.ToString()));
                    reply += "\n";
                }
            }
            if (reply != "")
                await ReplyAsync(reply);
            else
                await ReplyAsync("Sheet is empty.");
        }

        [Command("list")]
        [Alias("l")]
        [Summary("Lists all available sheets.")]
        public async Task ListSheets()
        {
            List<string> sheets = await signupHandler.SheetsList();
            if (sheets.Count == 0)
                await ReplyAsync("No available sheets.");
            else
            {
                // TO DO
                // make an embed for this
                await ReplyAsync(string.Join(", ", sheets));
            }
        }
    }
}
