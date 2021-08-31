using Genie, Genie.Router, Genie.Renderer.Json, Genie.Requests
using HTTP
import Genie.Router: route
import Genie.Renderer.Json: json
using ArgParse

# Add the command line argument to optionally specify the port desired
s = ArgParseSettings()
@add_arg_table s begin
    "--port"
        help = "an option with an argument to specify the Port number"
        arg_type = Int
end

# Parse arguments
parsed_args = parse_args(ARGS, s)
port_num = parsed_args["port"]

Genie.config.run_as_server = true

# Parse script: replace Julia code arguments with the actual value inputs from Tableau.
function parse_script(message)

  script = message["script"]
  num_args = length(message["data"])

  # Replace the script _arg* placeholders with the actual values. Return parsed script.
  for index in 1:num_args
    arg = "_arg" * string(index)
    arg_value = message["data"][arg]
    script = replace(script, arg => arg_value)
  end

  script

end

### Set up /info endpoint for Tableau connection.
route("/info") do
  json("""
  {"description": "", "creation_time": "0", "state_path": "", "server_version": "0.8.7", "name": "Julia", "versions": {"v1": {"features": {}}}}
  """)
end

### Evaluate the Julia script with the data provided by Tableau. Return output in JSON format.
route("/evaluate", method = POST) do

  # Parse json from Tableau into Julia Dictionary.
  message = jsonpayload()

  # Parse the script's arguments with their actual values.
  parsed_script = parse_script(message)

  # Evaluate the parsed script. Return in JSON format to Tableau for interpretability.
  output = eval(Meta.parse(parsed_script))
  json(output)

end

# Can change 8011 port to port of interest
Genie.startup(port_num, async = false)
