using Discord;
using Discord.Commands;
using Discord.WebSocket;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BasicBot.Modules
{
    public class Admin : ModuleBase<SocketCommandContext>
    {
        // Deletes a set amount of messages from the channel it was called in
        [Command("purge")]
        [Summary("Clears set number of messages from channel it's called in.")]
        [RequireBotPermission(GuildPermission.ManageMessages)] // Command will only execute if the bot has permissons to manage messages in server
        [RequireUserPermission(GuildPermission.ManageMessages)] // Command will only execute if the user has permissions to manage messages in server
        [Alias("delete", "clear")]
        public async Task Purge([Summary("How many messages to be deleted")] int number = 0)
        {
            // Discord limits how many messages can be bulk deleted at once
            if (number < 100)
            {
                try
                {
                    // Get messages to delete from channel command was used
                    var messagesToDelete = await (Context.Channel as SocketTextChannel).GetMessagesAsync(number + 1).FlattenAsync();
                    // Attempt to delete selected messages
                    await (Context.Channel as SocketTextChannel).DeleteMessagesAsync(messagesToDelete);
                    // Send response saying how many were deleted
                    await Context.Channel.SendMessageAsync($"{Context.User.Username} deleted {number} messages");
                    // Wait 10 seconds and delete message saying how many were deleted
                    var task = Task.Run(async () =>
                    {
                        var DelMsg = await (Context.Channel as SocketTextChannel).GetMessagesAsync(1).FlattenAsync();
                        await Task.Delay(10000);
                        await (Context.Channel as SocketTextChannel).DeleteMessagesAsync(DelMsg);
                    });
                }
                catch (Exception e)
                {
                    // Send error message
                    await ReplyAsync("Error: " + e.Message);
                }
            }
            else
            {
                await Context.Channel.SendMessageAsync("You cannot delete more than 99 messages");
            }
        }
    }
}
