-- About SQLQuerier.lua
--
-- Used for running SQL queries in the ILLiad client.

luanet.load_assembly("System");
luanet.load_assembly("System.Xml");
luanet.load_assembly("System.Windows.Forms");

local webClient = luanet.import_type("System.Net.WebClient");
local clipboard = luanet.import_type("System.Windows.Forms.Clipboard");

local interfaceMngr = nil;

local SQLForm = {};
SQLForm.Form = nil;
SQLForm.RibbonPage = nil;

function Init()
	interfaceMngr = GetInterfaceManager();
	
	SQLForm.Form = interfaceMngr:CreateForm("SQL Querier", "Script");
	
	SQLForm.RibbonPage = SQLForm.Form:CreateRibbonPage("SQL Querier");
	SQLForm.Browser = SQLForm.Form:CreateBrowser("SQL Querier", "SQL Browser", "Resource Browser");
	
	SQLForm.Query = SQLForm.Form:CreateMemoEdit("Query: ", "");

	SQLForm.Results = SQLForm.Form:CreateMemoEdit("Results (paste into Excel, and enjoy!): ", "");
	SQLForm.Results.ReadOnly = true;
	
	SQLForm.Num = SQLForm.Form:CreateTextEdit("Number of rows: " , "");
	SQLForm.Num.ReadOnly = true;
	
	SQLForm.RibbonPage:CreateButton("SQL Help", GetClientImage("Help32"), "W3Info", "Resources");
	SQLForm.RibbonPage:CreateButton("Database Tables", GetClientImage("About32"), "DataTables", "Resources");	
	
	SQLForm.RibbonPage:CreateButton("Clear Results", GetClientImage("Undo32"), "ClearResults", "Controls");
	SQLForm.RibbonPage:CreateButton("Copy Results", GetClientImage("Copy32"), "CopyResults", "Controls");
	SQLForm.RibbonPage:CreateButton("Run Query", GetClientImage("ExportData32"), "RunQuery", "Controls");
	
	SQLForm.Form:LoadLayout("layout.xml");
	SQLForm.Form:Show();
	
	SQLForm.Browser:Navigate("https://support.atlas-sys.com/hc/en-us/articles/360011812074-ILLiad-Database-Tables");
end

function RunQuery()
	SQLForm.Results.Value = "";
	SQLForm.Num.Value = "";
	
	local query = SQLForm.Query.Value;
	
	local results = PullData(query);
	if not results then
		interfaceMngr:ShowMessage("No results.", "Oh no!");
		return;
	end
	
	local report = "";
	
	for ct = 0, results.Columns.Count - 1 do
		report = report .. results.Columns:get_Item(ct).ColumnName .. "\t";
		
	end
	
	report = report .. "\n";
	
	for ct = 0, results.Rows.Count - 1 do
		
		for dt = 0, results.Columns.Count - 1 do
			report = report .. tostring(results.Rows:get_Item(ct):get_Item(dt)).. "\t";
		end
	
		report = report .. "\n";
	end
	
	SQLForm.Results.Value = report:gsub(": %d+", ""); -- Gsub gets rid of null values that appear as numbers.
	SQLForm.Num.Value = results.Rows.Count;
end

function CopyResults()
	clipboard.SetData("Text", SQLForm.Results.Value);
end

function ClearResults()
	SQLForm.Results.Value = "";
	SQLForm.Num.Value = "";
end

function W3Info()
	SQLForm.Browser:Navigate("https://www.w3schools.com/sql/");
end

function DataTables()
	SQLForm.Browser:Navigate("https://support.atlas-sys.com/hc/en-us/articles/360011812074-ILLiad-Database-Tables");
end

function PullData(query) -- Used for SQL queries that will return more than one result.
	local connection = CreateManagedDatabaseConnection();
	function PullData2()
		connection.QueryString = query;
		connection:Connect();
		local results = connection:Execute();
		connection:Disconnect();
		connection:Dispose();
		
		return results;
	end
	
	local success, results = pcall(PullData2, query);
	if not success then
		connection:Disconnect();
		connection:Dispose();
		return false;
	end
	
	return results;
end