% ======================
% Step 1: Device Energy Consumption Calculation
% ======================
fileName = 'Final.xlsx'; 
sheetDevices = 'Government Elec'; %sheet name for device data

% Import data from the specified sheet
data = readtable(fileName, 'Sheet', sheetDevices);

% Assuming the data contains the columns 'Device', 'ConsumptionW', and 'TimeWorkedH'
power_kW = data.ConsumptionW / 1000; % Convert power from watts to kilowatts
hours = data.TimeWorkedH; % Column for hours of operation

% Calculate energy consumption for each device (in kilowatt-hours)
data.Energy_kWh_PerDay = power_kW .* hours; % Energy consumed per day (in kWh)

% Calculate the total energy consumption per day in kWh
totalEnergyConsumption_PerDay_kWh = sum(data.Energy_kWh_PerDay);
% Calculate the total energy consumption per month in kWh (assuming 30 days in a month)
totalEnergyConsumption_PerMonth_kWh = totalEnergyConsumption_PerDay_kWh * 30;

% Display results: Show organized table with device names and energy consumption
disp('Device Energy Consumption Table (in kWh):');
disp(table(data.Device, power_kW, data.TimeWorkedH, data.Energy_kWh_PerDay, ...
    'VariableNames', {'Device', 'Power_kW', 'TimeWorkedH', 'Energy_kWh_PerDay'}));

% Display total energy consumption
fprintf('Total Energy Consumption Per Month: %.2f kWh\n', totalEnergyConsumption_PerMonth_kWh);

% ======================
% Step 2: Solar Energy Production Calculation
% ======================
sheetSolar = 'Solar';

% Import data from the specified sheet for solar energy production
dataSolar = readtable(fileName, 'Sheet', sheetSolar);

% Ask user for solar panel details
numSolarPanels = input('Enter the number of solar panels: ');
solarPanelProductionPerHour_kW = input('Enter the production per solar panel per hour in kW: ');

% Calculate total production per hour based on user input (in kW)
ProductionPerHour_kW = numSolarPanels * solarPanelProductionPerHour_kW; % Total production in kW

% Add a new column for ProductionPerHour to the dataSolar table (same value for all rows)minSolarDayIdx
dataSolar.ProductionPerHour_kW = repmat(ProductionPerHour_kW, height(dataSolar), 1);

% Assuming the data contains the columns: 'Days', 'SunHours'
days = dataSolar.Days; % Number of days
sunHours = dataSolar.SunHours; % Hours of sun per day

% Calculate total solar production per day in kilowatt-hours (total production in kW * hours of sun per day)
dataSolar.Total_Production_kWh_PerDay = sunHours .* dataSolar.ProductionPerHour_kW; % Total in kWh per day

% Calculate total solar production per month (in kWh)
totalSolarProduction_PerMonth_kWh = sum(dataSolar.Total_Production_kWh_PerDay);

% Display results for solar production
disp('Solar Energy Production Table (in kWh):');
disp(table(dataSolar.Days, dataSolar.SunHours, dataSolar.ProductionPerHour_kW, dataSolar.Total_Production_kWh_PerDay, ...
    'VariableNames', {'Days', 'SunHours', 'ProductionPerHour_kW', 'Total_Production_kWh_PerDay'}));


% ======================
% Step 3: Calculate the Difference
% ======================
energyDifference_kWh = totalSolarProduction_PerMonth_kWh - totalEnergyConsumption_PerMonth_kWh; % Difference in kWh

% Display energy balance summary
fprintf('Total Device Energy Consumption Per Month: %.2f kWh\n', totalEnergyConsumption_PerMonth_kWh);
fprintf('Total Solar Energy Production Per Month: %.2f kWh\n', totalSolarProduction_PerMonth_kWh);
fprintf('Energy Difference (Production - Consumption): %.2f kWh\n', energyDifference_kWh);

% ======================
% Step 4: Handle Energy Deficit
% ======================
energyDeficit_kWh=abs(energyDifference_kWh);

if energyDifference_kWh < 0
    % Calculate the absolute value of the energy difference in kWh
    energyDeficit_kWh = abs(energyDifference_kWh);

    % Calculate the total sun hours for the entire period from the 'SunHours' column
    totalSunHours = sum(sunHours); % Use the sum of SunHours directly

    % Calculate the total production per panel for the entire period (in kWh)
    productionPerPanelTotal_kWh = solarPanelProductionPerHour_kW * totalSunHours; % Total production using total SunHours

    % Calculate the number of additional panels needed (rounding up to the next whole panel)
    additionalPanels = ceil(energyDeficit_kWh / productionPerPanelTotal_kWh);

    % Display the number of additional panels required
    fprintf('To compensate the energy difference, you need an additional %d solar panels.\n', additionalPanels);
    

% ======================
% Step 5: Calculate Electricity Cost Based on Tariff
% ======================

% Determine the tariff category and calculate the cost
if energyDeficit_kWh <= 50
    tariffCategory = 'الأولى';
    costPerKWh = 0.68; 
elseif energyDeficit_kWh <= 100
    tariffCategory = 'الثانية';
    costPerKWh = 0.78; 
elseif energyDeficit_kWh <= 200
    tariffCategory = 'الثالثة';
    costPerKWh = 0.95; 
elseif energyDeficit_kWh <= 350
    tariffCategory = 'الرابعة';
    costPerKWh = 1.55;
elseif energyDeficit_kWh <= 650
     tariffCategory = 'الخامسة';
     costPerKWh = 1.95;
elseif energyDeficit_kWh <= 1000
     tariffCategory = 'السادسة';
     costPerKWh = 2.1; 
else
    tariffCategory = 'السابعة';
    costPerKWh = 2.2; 
end

% Calculate the cost
monthlyCost = energyDeficit_kWh * costPerKWh; 

% Display the deficit, tariff category, and cost
fprintf('Energy Deficit: %.2f kWh\n', energyDeficit_kWh);
fprintf('Tariff Category: %s\n', tariffCategory);
fprintf('Monthly Electricity Cost: %.2f EGP\n', monthlyCost);

elseif energyDifference_kWh == 0
    % Case when energy consumption equals energy production
    fprintf('Your energy consumption equals your energy production. No additional panels are needed.\n');
    
else
    % Case when energy production is greater than consumption
    fprintf('Energy production is sufficient! No additional cost is required.\n');
end



% ======================
% Step 6: Plot the Data (with subplots)
% ======================


% ==========================================
% 1. Plot Daily Solar Energy Production (kWh)
% ==========================================
subplot(2, 2, 1); % First subplot (2 rows, 2 columns, position 1)
plot(days, dataSolar.Total_Production_kWh_PerDay, 'b-', 'LineWidth', 2);
title('Daily Solar Energy Production (kWh)');
xlabel('Day of the Month');
ylabel('Energy Produced (kWh)');
grid on;

% ==========================================
% 2. Calculate and Plot Energy Deficit (kWh)
% ==========================================
% Calculate the daily energy deficit (production - consumption)
energyDeficitDaily_kWh = dataSolar.Total_Production_kWh_PerDay - sum(data.Energy_kWh_PerDay);

% Create the area chart for energy deficit
subplot(2, 2, 2); % Second subplot (2 rows, 2 columns, position 2)
area(days, energyDeficitDaily_kWh, 'FaceColor', 'm'); % Area chart for the deficit over days
title('Energy Deficit Over Days (kWh)');
xlabel('Day of the Month');
ylabel('Energy Deficit (kWh)');
grid on;

% ==========================================
% 3. Solar Energy Production vs. Sun Hours
% ==========================================
x = dataSolar.SunHours; % Sun hours per day
y = dataSolar.Total_Production_kWh_PerDay; % Daily solar energy production

subplot(2, 2, 3); % Third subplot
scatter(x, y, 'filled', 'MarkerFaceColor', 'b'); % Scatter plot
hold on;

% Linear regression and plotting the line
plot(x, polyval(polyfit(x, y, 1), x), 'r-', 'LineWidth', 2);

%polyfit بيختار أفضل معادلة لعمل فيتيننج للمتغيرات
%polyval  يحسب القيم المتوقعة عن طريق الميل من المعادلة
% Labels and title
title('Solar Energy Production vs. Sun Hours');
xlabel('Sun Hours per Day');
ylabel('Solar Energy Production (kWh)');
grid on;
hold off;

% ==========================================
% 4. Plot Device Energy Consumption per Device
% ==========================================
subplot(2, 2, 4); % Fourth subplot (2 rows, 2 columns, position 4)
deviceIndex = 1:height(data); % Numeric index for devices
bar(deviceIndex, data.Energy_kWh_PerDay, 'FaceColor', 'c'); % Bar chart for energy consumed
set(gca, 'xtick', deviceIndex, 'xticklabel', data.Device); % Set x-axis labels to device names
%get current axes 
title('Energy Consumption per Device (kWh per Day)');
xlabel('Device');
ylabel('Energy Consumed (kWh)');
xtickangle(45); % Rotate x-axis labels for better visibility
grid on;

% ======================
% Step7: Additional Calculations
% ======================

% 1. Find the device with maximum energy consumption (in kWh)
[maxDeviceConsumption, maxDeviceIdx] = max(data.Energy_kWh_PerDay);
[minDeviceConsumption, minDeviceIdx] = min(data.Energy_kWh_PerDay);

% 2. Find the day with maximum solar energy production (in kWh)
[maxSolarProduction, maxSolarDayIdx] = max(dataSolar.Total_Production_kWh_PerDay);
[minSolarProduction, minSolarDayIdx] = min(dataSolar.Total_Production_kWh_PerDay);

% 3. Calculate the average energy consumption and solar production
avgConsumptionPerDay = mean(data.Energy_kWh_PerDay); % Average consumption per day 
avgSolarProduction = mean(dataSolar.Total_Production_kWh_PerDay);

% 4. Calculate the standard deviation for consumption and production
stdDeviceConsumption = std(data.Energy_kWh_PerDay);
stdSolarProduction = std(dataSolar.Total_Production_kWh_PerDay);

% ======================
% Step 8: Display Results
% ======================

% Display device with maximum and minimum consumption
fprintf('Device with Maximum Consumption: %s with %.2f kWh\n', data.Device{maxDeviceIdx}, maxDeviceConsumption);
fprintf('Device with Minimum Consumption: %s with %.2f kWh\n', data.Device{minDeviceIdx}, minDeviceConsumption);


% Display the day with maximum solar energy production
fprintf('Day with Maximum Solar Energy Production: Day %d with %.2f kWh\n', maxSolarDayIdx, maxSolarProduction);
fprintf('Day with Minimum Solar Energy Production: Day %d with %.2f kWh\n', minSolarDayIdx, minSolarProduction);


% Display the averages and standard deviations
fprintf('Average Consumption per day: %.2f kWh\n', avgConsumptionPerDay); % Changed output
fprintf('Average Solar Production per day: %.2f kWh\n', avgSolarProduction);
fprintf('Standard Deviation of Device Consumption: %.2f kWh\n', stdDeviceConsumption);
fprintf('Standard Deviation of Solar Production: %.2f kWh\n', stdSolarProduction);


