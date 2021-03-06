using BusinessLogic;
using Discord;
using Discord.Commands;
using Discord.WebSocket;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace BasicBot.Modules
{
    public class General : ModuleBase<SocketCommandContext>
    {
        // Command service is used here to get command info for the help command
        private CommandService _service;
        private MiscApiHandler _apiHandler = new MiscApiHandler();

        // Constructor to inject command service
        public General(CommandService service)
        {
            _service = service;
        }

        [Command("help")] // Name of command, how it will be called in discord, ex: "!help"
        [Alias("h", "halp")] // Alternate names that can call this command, separate aliases with a comma to list more than one
        [Summary("Shows what commands are available, or will return more information for a specific command.")] // Description of what the command does, used for help information
        // Command functions MUST be async
        public async Task HelpAsync([Remainder, Summary("Command to retrieve help for")] string command = null)
        // Remainder will accept all remaining chars as parameter, can only be used as last parameter
        // Summary describes the parameter that is accepted
        {
            // Prefix to show in embed with help
            string prefix = "!";
            // No command is given, return list of commands
            if (command == null)
            {
                // Create embed object to add to
                // use an embed visualizer to design and get code for the embed
                // https://leovoel.github.io/embed-visualizer/
                var builder = new EmbedBuilder()
                {
                    Color = new Color(114, 137, 218),
                    Description = $"**These are the commands available**"
                };
                // Loop through modules available
                foreach (var module in _service.Modules)
                {
                    string description = null;
                    // Sort commands alphabetically
                    var sorted = module.Commands.OrderBy(x => x.Name);
                    // Loop through sorted commands
                    foreach (var cmd in sorted)
                    {
                        // Check if commands pass preconditions
                        var result = await cmd.CheckPreconditionsAsync(Context);
                        // If passed, add command's name to list of commands
                        if (result.IsSuccess)
                            description += $"{prefix}{cmd.Aliases.First()}\n";
                    }
                    // if the module wasn't empty, add the commands to the embed
                    if (!string.IsNullOrWhiteSpace(description))
                    {
                        builder.AddField(x =>
                        {
                            // module.Name is the name of the class in the modules folder that contains the command functions, this will be shown in the embed as a category of commands
                            x.Name = module.Name;
                            // description is the list of commands created when looping through modules
                            x.Value = description;
                            x.IsInline = false;
                        });
                    }
                }

                // These lines will send the help command embed to the User's private messages, can use to send a message to the user instead of the channel, reduces chat bloat
                //var dmchannel = await Context.User.GetOrCreateDMChannelAsync();
                //await dmchannel.SendMessageAsync("", false, builder.Build());

                // This will send help embed to channel the command was used in
                await ReplyAsync("", false, builder.Build());
            }
            // A specific command was requested information on
            else
            {
                // Search for matching command
                var result = _service.Search(Context, command);
                // No matching command was found
                if (!result.IsSuccess)
                {
                    await ReplyAsync($"Sorry, I couldn't find a command like **{command}**.");
                    return;
                }
                // Create an embed object to send back
                var builder = new EmbedBuilder()
                {
                    Color = new Color(114, 137, 218),
                    Description = $"**Help for command {prefix}{command}**\n"
                };

                foreach (var match in result.Commands)
                {
                    var cmd = match.Command;
                    // Add command info to the embed
                    builder.AddField(x =>
                    {
                        x.Name = "**Aliases:**";
                        // Combines list of aliases into a string to add to the embed
                        x.Value += string.Join(", ", cmd.Aliases);
                        x.Value += $"\n**Summary:** \n{cmd.Summary}";
                        // If there are parameters, create a section in the embed for them
                        if (cmd.Parameters.Select(p => p.Name).FirstOrDefault() != null)
                            x.Value += $"\n**Parameters:**\n";
                        // Loop and add all parameters and summaries
                        foreach (var param in cmd.Parameters)
                        {
                            x.Value += $"{param.Name}: {param.Summary}\n";
                        }
                        // Add a usage example for even more clarity
                        x.Value += $"\n**Usage:** \n{prefix}{command} ";
                        // If command has parameters, add them to the usage
                        if (cmd.Parameters.Select(p => p.Name).FirstOrDefault() != null)
                        {
                            x.Value += string.Join(" ", cmd.Parameters);
                        }
                        x.IsInline = false;
                    });

                }
                // Send embed back to channel command was used in
                await ReplyAsync("", false, builder.Build());
            }
        }

        // Simple Hello command
        [Command("hello")]
        [Alias("hi")]
        [Summary("Say hello!")]
        public async Task Hello()
        {
            await Context.Channel.SendMessageAsync("no u");
        }

        // Adds or removes a role to a user, must have admin permissions to add roles that have admin access
        [Command("role")]
        [Alias("r")]
        [Summary("Adds or removes a discord role.")]
        public async Task Role([Remainder, Summary("Name of the role you wish to add to your account.")] string role)
        {
            try
            {
                // Get variables needed to add roles
                SocketGuild guild = Context.Guild;
                SocketGuildUser user = guild.Users.Where(x => x.Id == Context.User.Id).FirstOrDefault();
                SocketRole targetRole = guild.Roles.Where(x => x.Name.ToUpper() == role.ToUpper()).FirstOrDefault();
                // Return from command if the role does not exist
                if (targetRole == null)
                {
                    await ReplyAsync($"Error: role `{role}` does not exist.");
                    return;
                }
                // Return from the command if the user attempts to add an admin role without proper access
                if (targetRole.Permissions.Administrator == true && user.GuildPermissions.Administrator == false)
                {
                    await ReplyAsync("You do not have permissions to add this role.");
                    return;
                }
                // User already has the role, remove it from their role list
                if (user.Roles.Contains(targetRole))
                {
                    await user.RemoveRoleAsync(targetRole);
                    await ReplyAsync($"Removed {targetRole.Name} role.");
                }
                // User does not have the role already, add it to their list
                else
                {
                    await user.AddRoleAsync(targetRole);
                    await ReplyAsync($"Added {targetRole.Name} role");
                }

            }
            catch (Exception e)
            {
                // Freak chance that something broke
                Console.WriteLine(e);
                await ReplyAsync("Error: Something went wrong.");
            }
        }

        [Command("cat")]
        [Alias("kitty", "meow", "randomcat", "randomkitty")]
        [Summary("Random cat images. Because reasons.")]
        public async Task RandomCat()
        {
            string cat = await _apiHandler.RandomCat();
            if (!string.IsNullOrWhiteSpace(cat))
                await ReplyAsync(cat);
            else
                await ReplyAsync("An error occurred.");
        }
    }
}
