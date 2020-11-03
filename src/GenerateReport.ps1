$file = Get-Content $args[0]
$containsSemiColumn = $file | %{$_ -match ";"}
$wrongHeader = $file | %{$_ -match "Column1"}
if ($containsSemiColumn -contains $true) {
    if ($wrongHeader -contains $true) {
        $database = Get-Content -Path $args[0] | Select-Object -Skip 1 | ConvertFrom-Csv -Delimiter ';'  
    } else {
        $database = Import-Csv -Path $args[0] -Delimiter ';'
    }
} else {
    if ($wrongHeader -contains $true) {
        $database = Get-Content -Path $args[0] | Select-Object -Skip 1 | ConvertFrom-Csv -Delimiter ','  
    } else {
        $database = Import-Csv -Path $args[0] -Delimiter ','
    }
}
$database = $database | ConvertTo-Html -Fragment

$head = @"
<title></title>
<link href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.0/css/bootstrap.min.css" rel="stylesheet" id="bootstrap-css">
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.0/js/bootstrap.min.js"></script>
<script src="https://code.jquery.com/jquery-1.11.1.min.js"></script>
<script src='https://www.chartjs.org/dist/2.9.3/Chart.min.js'></script>
<script src='https://www.chartjs.org/samples/latest/utils.js'></script>
<script src="https://cdn.jsdelivr.net/gh/emn178/chartjs-plugin-labels/src/chartjs-plugin-labels.js"></script>
<link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css" integrity="sha384-JcKb8q3iqJ61gNV9KGb8thSsNjpSL0n8PARn9HuZOnIxN0hoP+VmmDGMN5t9UJ0Z" crossorigin="anonymous">
<script src="https://code.jquery.com/jquery-3.5.1.slim.min.js" integrity="sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj" crossorigin="anonymous"></script>
<script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.1/dist/umd/popper.min.js" integrity="sha384-9/reFTGAW83EW2RDu2S0VKaIzap3H66lZH81PoYlFhbGU+6BZp6G7niu735Sk7lN" crossorigin="anonymous"></script>
<script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js" integrity="sha384-B4gt1jrGC7Jh4AgTPSdUtOBvfO8shuf57BaghqFfPlYxofvL8/KUEfYiJOMMV+rV" crossorigin="anonymous"></script>
<style>
    li a {
        color: white;
        background-color: rgba(0,0,0,0.5);
        font-size: large;
    }
    div p {
        color: white;
    }
</style>
<!------ Include the above in your HEAD tag ---------->

<div id="wrapper">
<div class="overlay"></div>
<!-- Sidebar -->
        <!-- /#sidebar-wrapper -->
    </div>
    <!-- /#wrapper -->
<style>
    canvas {
        -moz-user-select: none;
        -webkit-user-select: none;
        -ms-user-select: none;
    }
</style>
"@

$postContent = @"
<div class="container">
        <div class="container" style="width:75%;">
            <canvas id="canvas-total-capacity"></canvas>
        </div>
        <br>
        <hr>
        <br>
        <div class="container" style="width:75%;">
            <canvas id="canvas-main-chart-doughnut"></canvas>
        </div>
        <br>
        <hr>
        <br>
        <div class="container" style="width:75%;">
            <canvas id="canvas-main-chart-pie"></canvas>
        </div>
        <br>
        <hr>
        <br>
        <div class="container" style="width:75%;">
            <canvas id="canvas-main-chart-polar"></canvas>
        </div>
        <br>
        <hr>
        <br>
        <div class="container" style="width:75%;">
            <canvas id="canvas-main-chart-radar"></canvas>
        </div>
        <br>
        <hr>
        <br>
        <div class="container" style="width:75%;">
            <canvas id="canvas-main-chart-basic-line"></canvas>
        </div>
        <br>
        <hr>
        <br>
        <div class="container" style="width:75%;">
            <canvas id="canvas-main-chart-stacked-bar"></canvas>
        </div>
        <br>
        <hr>
        <br>
        <div class="container" style="width:75%;">
            <canvas id="canvas-main-chart-stacked-line"></canvas>
        </div>
        <br>
        <hr>
        <br>
        <div class="container" style="width:75%;">
            <canvas id="canvas-main-chart-vertical-bar"></canvas>
        </div>
        <br>
        <hr>
        <br>
        <div class="container" style="width:75%;">
            <canvas id="canvas-main-chart-horizontal-bar"></canvas>
        </div>
        <br>
        <hr>
        <br>
        <div class="container" style="width:75%;">
            <canvas id="canvas-free-space"></canvas>
        </div>
    </div>

	<script>

    var color = Chart.helpers.color;

    function getRandomColor() {
        var letters = '0123456789ABCDEF';
        var color = '#';
        for (var i = 0; i < 6; i++) {
            color += letters[Math.floor(Math.random() * 16)];
        }
        return color;
    }

    function createStackedLineChart(labels, dataRows, dataNames, borderColor, backgroundColor, title, xlabel, ylabel) {

        var datasets = [];
        for (let index = 0; index < dataRows.length; index++) {
            var dataset = {};
            dataset.label = dataNames[index];
            dataset.data = dataRows[index];
            dataset.borderColor =  borderColor[index];
		    dataset.backgroundColor = backgroundColor[index];
            datasets.push(dataset);
        }

		var config = {
			type: 'line',
			data: {
				labels: labels,
				datasets: datasets
			},
			options: {
				responsive: true,
				title: {
					display: true,
					text: title,
                    fontSize: 36
				},
				tooltips: {
                    mode: 'label',
                    intersect: false,
                    callbacks: {
                        afterTitle: function() {
                            window.total = 0;
                        },
                        label: function(tooltipItem, data) {
                            var efs = data.datasets[tooltipItem.datasetIndex].label;
                            var space = data.datasets[tooltipItem.datasetIndex].data[tooltipItem.index];
                            window.total += space;
                            return efs + ": " + space.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");             
                        },
                        footer: function() {
                            return "TOTAL: " + window.total.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");
                        }
                    }
                },
				hover: {
					mode: 'index'
				},
				scales: {
					xAxes: [{
						scaleLabel: {
							display: true,
							labelString: xlabel
						}
					}],
					yAxes: [{
						stacked: true,
						scaleLabel: {
							display: true,
							labelString: ylabel
						}
					}]
				}
			}
		};
        return config;
    }

    function createBasicLineChart(labels, dataRows, dataNames, borderColor, backgroundColor, title, xlabel, ylabel) {

        var datasets = [];
        for (let index = 0; index < dataRows.length; index++) {
            var dataset = {};
            dataset.label = dataNames[index];
            dataset.data = dataRows[index];
            dataset.fill = false;
            dataset.borderColor =  borderColor[index];
		    dataset.backgroundColor = backgroundColor[index];
            datasets.push(dataset);
        }

        var config = {
            type: 'line',
            data: {
                labels: labels,
                datasets: datasets
            },
            options: {
                responsive: true,
                title: {
                    display: true,
                    text: title,
                    fontSize: 36
                },
                tooltips: {
                    mode: 'label',
                    intersect: false,
                    callbacks: {
                        afterTitle: function() {
                            window.total = 0;
                        },
                        label: function(tooltipItem, data) {
                            var efs = data.datasets[tooltipItem.datasetIndex].label;
                            var space = data.datasets[tooltipItem.datasetIndex].data[tooltipItem.index];
                            window.total += space;
                            return efs + ": " + space.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");             
                        },
                        footer: function() {
                            return "TOTAL: " + window.total.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");
                        }
                    }
                },
                hover: {
                    mode: 'index'
                },
                scales: {
                    xAxes: [{
                        scaleLabel: {
                            display: true,
                            labelString: xlabel
                        }
                    }],
                    yAxes: [{
                        scaleLabel: {
                            display: true,
                            labelString: ylabel
                        }
                    }]
                }
            }
        };
        return config;
    }

    function createTowerChart(labels, totalCapacity, dataNames, borderColor, backgroundColor, title, xlabel, ylabel) {

        var datasets = [];
        for (let index = 0; index < totalCapacity.length; index++) {
            var dataset = {};
            dataset.label = dataNames[index];
            dataset.data = [totalCapacity[index]];
            dataset.borderColor =  borderColor[index];
            dataset.backgroundColor = backgroundColor[index];
            datasets.push(dataset);
        }

        var config = {
            type: 'bar',
				data: {
                    labels: labels,
			        datasets: datasets
                },
				options: {
					title: {
						display: true,
						text: title,
                        fontSize: 36
					},
                    plugins: {
                    labels: {
                        render: () => {}
                    }
                    },
					tooltips: {
						mode: 'index',
						intersect: false
					},
					responsive: true,
					scales: {
						xAxes: [{
							stacked: true,
                            scaleLabel: {
                                display: true,
                                labelString: xlabel
                            }
						}],
						yAxes: [{
							stacked: true,
                            scaleLabel: {
                                display: true,
                                labelString: ylabel
                            }
						}]
					}
				}
        };
        return config;
    }

    function createStackedBarChart(labels, dataRows, dataNames, borderColor, backgroundColor, title, xlabel, ylabel) {

        var datasets = [];
        for (let index = 0; index < dataRows.length; index++) {
            var dataset = {};
            dataset.label = dataNames[index];
            dataset.data = dataRows[index];
            dataset.borderColor =  borderColor[index];
            dataset.backgroundColor = backgroundColor[index];
            datasets.push(dataset);
        }

        var config = {
            type: 'bar',
            data: {
                labels: labels,
                datasets: datasets
            },
            options: {
                responsive: true,
                title: {
                    display: true,
                    text: title,
                    fontSize: 36
                },
                plugins: {
                  labels: {
                    render: () => {}
                  }
                },
                tooltips: {
                    mode: 'label',
                    intersect: false,
                    callbacks: {
                        afterTitle: function() {
                            window.total = 0;
                        },
                        label: function(tooltipItem, data) {
                            var efs = data.datasets[tooltipItem.datasetIndex].label;
                            var space = data.datasets[tooltipItem.datasetIndex].data[tooltipItem.index];
                            window.total += space;
                            return efs + ": " + space.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");             
                        },
                        footer: function() {
                            return "TOTAL: " + window.total.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");
                        }
                    }
                },
                hover: {
                    mode: 'index'
                },
                scales: {
                    xAxes: [{
                        scaleLabel: {
                            display: true,
                            labelString: xlabel
                        },
                        stacked: true
                    }],
                    yAxes: [{
                        scaleLabel: {
                            display: true,
                            labelString: ylabel
                        },
                        stacked: true
                    }]
                }
            }
        };
        return config;
    }

    function createRadarChart(labels, dataRows, dataNames, borderColor, backgroundColor, title, xlabel, ylabel) {

        var labels = dataNames.slice(0,dataNames.length-1)

        var dataArray = [];
        for (let index = 0; index < totalCapacity.length; index++) {
            dataArray.push(totalCapacity[index])
        }

        var datasets = [{
            data: dataArray,
            backgroundColor: backgroundColor[1],
            borderColor: borderColor[1],
            label: title
        }];

        var config = {
            type: 'radar',
            data: {
                labels: labels,
                datasets: datasets
            },
            options: {
                responsive: true,
                title: {
                    display: true,
                    text: title,
                    fontSize: 36
                },
                hover: {
                    mode: 'index'
                }
            }
        };
        return config;
    }

    function createDoughnutChart(totalCapacity, dataNames, borderColor, backgroundColor, title, xlabel, ylabel) {

        var labels = dataNames.slice(0,dataNames.length-1)

        var dataArray = [];
        for (let index = 0; index < totalCapacity.length; index++) {
            dataArray.push(totalCapacity[index])
        }

        var datasets = [{
            data: dataArray,
            backgroundColor: backgroundColor,
        }];

        var config = {
            type: 'doughnut',
            data: {
                labels: labels,
                datasets: datasets
            },
            options: {
                responsive: true,
                title: {
                    display: true,
                    text: title,
                    fontSize: 36
                },
                hover: {
                    mode: 'index'
                }
            }
        };
        return config;
    }

    function createPieChart(labels, dataRows, dataNames, borderColor, backgroundColor, title, xlabel, ylabel) {

        var labels = dataNames.slice(0,dataNames.length-1)

        var dataArray = [];
        for (let index = 0; index < totalCapacity.length; index++) {
            dataArray.push(totalCapacity[index])
        }

        var datasets = [{
            data: dataArray,
            backgroundColor: backgroundColor,
        }];

        var config = {
            type: 'pie',
            data: {
                labels: labels,
                datasets: datasets
            },
            options: {
                responsive: true,
                title: {
                    display: true,
                    text: title,
                    fontSize: 36
                },
                hover: {
                    mode: 'index'
                }
            }
        };
        return config;
    }

    function createPolarChart(labels, dataRows, dataNames, borderColor, backgroundColor, title, xlabel, ylabel) {

        var labels = dataNames.slice(0,dataNames.length-1)

        var dataArray = [];
        for (let index = 0; index < totalCapacity.length; index++) {
            dataArray.push(totalCapacity[index])
        }

        var datasets = [{
            data: dataArray,
            backgroundColor: backgroundColor,
        }];

        var config = {
            data: {
                labels: labels,
                datasets: datasets
            },
            options: {
                responsive: true,
                title: {
                    display: true,
                    text: title,
                    fontSize: 36
                },
                plugins: {
                    labels: {
                        render: () => {}
                    }
                    },
                hover: {
                    mode: 'index'
                }
            }
        };
        return config;
    }

    function createVerticalBarChart(labels, dataRows, dataNames, borderColor, backgroundColor, title, xlabel, ylabel) {

        var datasets = [];
        for (let index = 0; index < dataRows.length; index++) {
            var dataset = {};
            dataset.label = dataNames[index];
            dataset.data = dataRows[index];
            dataset.borderColor =  borderColor[index];
            dataset.backgroundColor = backgroundColor[index];
            datasets.push(dataset);
        }

        var config = {
            type: 'bar',
            data: {
                labels: labels,
                datasets: datasets
            },
            options: {
                responsive: true,
                title: {
                    display: true,
                    text: title,
                    fontSize: 36
                },
                plugins: {
                  labels: {
                    render: () => {}
                  }
                },
                tooltips: {
                    mode: 'label',
                    intersect: false,
                    callbacks: {
                        afterTitle: function() {
                            window.total = 0;
                        },
                        label: function(tooltipItem, data) {
                            var efs = data.datasets[tooltipItem.datasetIndex].label;
                            var space = data.datasets[tooltipItem.datasetIndex].data[tooltipItem.index];
                            window.total += space;
                            return efs + ": " + space.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");             
                        },
                        footer: function() {
                            return "TOTAL: " + window.total.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");
                        }
                    }
                },
                scales: {
                    xAxes: [{
                        scaleLabel: {
                            display: true,
                            labelString: xlabel
                        },
                    }],
                    yAxes: [{
                        scaleLabel: {
                            display: true,
                            labelString: ylabel
                        },
                    }]
                }
            }
        };
        return config;
    }

    function createHorizontalBarChart(labels, dataRows, dataNames, borderColor, backgroundColor, title, xlabel, ylabel) {

        var datasets = [];
        for (let index = 0; index < dataRows.length; index++) {
            var dataset = {};
            dataset.label = dataNames[index];
            dataset.data = dataRows[index];
            dataset.borderColor =  borderColor[index];
            dataset.backgroundColor = backgroundColor[index];
            datasets.push(dataset);
        }

        var config = {
            type: 'horizontalBar',
            data: {
                labels: labels,
                datasets: datasets
            },
            options: {
                responsive: true,
                title: {
                    display: true,
                    text: title,
                    fontSize: 36
                },
                legend: {
                    position: 'right',
				},
                tooltips: {
                    mode: 'label',
                    intersect: false,
                    callbacks: {
                        afterTitle: function() {
                            window.total = 0;
                        },
                        label: function(tooltipItem, data) {
                            var efs = data.datasets[tooltipItem.datasetIndex].label;
                            var space = data.datasets[tooltipItem.datasetIndex].data[tooltipItem.index];
                            window.total += space;
                            return efs + ": " + space.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");             
                        },
                        footer: function() {
                            return "TOTAL: " + window.total.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");
                        }
                    }
                },
                hover: {
                    mode: 'index'
                },
                scales: {
                    xAxes: [{
                        scaleLabel: {
                            display: true,
                            labelString: xlabel
                        },
                        stacked: true
                    }],
                    yAxes: [{
                        scaleLabel: {
                            display: true,
                            labelString: ylabel
                        },
                        stacked: true
                    }]
                }
            }
        };
        return config;
    }

    window.onload = function() {

        var table = document.getElementsByTagName("table")[0];
        table.style.display = 'none';

        labels = [];
        dataArray = [];
        dataNames = [];
        totalCapacity = [];
        efsFree = [];
        colorPalette = [
            'rgb(255, 99, 132)',
            'rgb(255, 159, 64)',
            'rgb(255, 205, 86)',
            'rgb(75, 192, 192)',
            'rgb(54, 162, 235)',
            'rgb(153, 102, 255)',
            'rgb(128, 128, 128)',
        ]

        nbOriginalColors = colorPalette.length;

        // Color Generator
        for (let k = 1; k < 5; k++) {
            // Darker shades
            for (let c = 0; c < nbOriginalColors; c++) {
                const color = colorPalette[c];
                var r = (parseInt(color.split('(')[1].split(',')[0]) - 40*k).toString();
                var g = (parseInt(color.split(', ')[1].split(',')[0]) - 40*k).toString();
                var b = (parseInt(color.split(', ')[2].split(')')[0]) - 40*k).toString();
                r = r < 0 ? 0 : r
                g = g < 0 ? 0 : g
                b = b < 0 ? 0 : b
                r = r > 255 ? 255 : r;
                g = g > 255 ? 255 : g;
                b = b > 255 ? 255 : b;
                var shade = 'rgb(' + r + ', ' + g + ', ' + b + ')';
                colorPalette.push(shade);
            }
            // Lighter shades
            for (let c = 0; c < nbOriginalColors; c++) {
                const color = colorPalette[c];
                var r = color.split('(')[1].split(',')[0];
                var g = (parseInt(color.split(', ')[1].split(',')[0]) + 40*k).toString();
                var b = (parseInt(color.split(', ')[2].split(')')[0]) + 40*k).toString();
                r = r < 0 ? 0 : r;
                g = g < 0 ? 0 : g;
                b = b < 0 ? 0 : b;
                r = r > 255 ? 255 : r;
                g = g > 255 ? 255 : g;
                b = b > 255 ? 255 : b;
                var shade = 'rgb(' + r + ', ' + g + ', ' + b + ')';
                colorPalette.push(shade);
            }
        }
        
        colorPaletteShallow = []

        colorPalette.forEach(col => {
            colorPaletteShallow.push(color(col).alpha(0.7).rgbString());
        });

        // Read header row
        rows = table.getElementsByTagName('tr');
        labelsRow = rows.item(0).getElementsByTagName('th');
        for (let index = 0; index < labelsRow.length ; index++) {
            labels[index] = labelsRow.item(index).innerText;
        }
        // Ignore first three columns
        labels = labels.slice(4);

        // Read data rows
        for (let index = 1; index < rows.length; index++) {
            const row = rows.item(index).getElementsByTagName('td');
            let tmpArray = [];
            let tmpFreeArray = [];
            dataNames[index-1] = row.item(2).innerText;
            totalCapacity[index-1] = parseFloat(row.item(3).innerText.replace(',','.'));
            for (let j = 4; j < row.length; j++) {
                tmpArray[j-4] = parseFloat(row.item(j).innerText.replace(',','.')) || 0;
                tmpFreeArray[j-4] = totalCapacity[index-1] - tmpArray[j-4];
            }
            dataArray.push(tmpArray);
            efsFree.push(tmpFreeArray);
        }

        // Read label column
        var documentTitle = rows.item(0).getElementsByTagName('th').item(1).innerText;
        var capacityLabel = rows.item(0).getElementsByTagName('th').item(3).innerText;
        var chartTitle = rows.item(1).getElementsByTagName('td').item(1).innerText;
        var yLabel = rows.item(2).getElementsByTagName('td').item(1).innerText;
        var xLabel = rows.item(3).getElementsByTagName('td').item(1).innerText;

        // Set document title
        document.getElementsByTagName('title').item(0).innerText = documentTitle;

        totalCapacity.pop();
        const total = dataArray.pop();
        const totalFree = efsFree.pop();

        var ctx = document.getElementById('canvas-free-space').getContext('2d');
        window.myLine = new Chart(ctx, createStackedLineChart(labels, efsFree, dataNames, colorPalette, colorPalette, "Free Space", xLabel, yLabel));

        var ctx = document.getElementById('canvas-total-capacity').getContext('2d');
        window.myLine = new Chart(ctx, createTowerChart([capacityLabel], totalCapacity, dataNames, colorPalette, colorPaletteShallow, capacityLabel, xLabel, yLabel));

        var ctx = document.getElementById('canvas-main-chart-basic-line').getContext('2d');
        window.myLine = new Chart(ctx, createBasicLineChart(labels, dataArray, dataNames, colorPalette, colorPaletteShallow, chartTitle, xLabel, yLabel));

        var ctx = document.getElementById('canvas-main-chart-stacked-line').getContext('2d');
        window.myLine = new Chart(ctx, createStackedLineChart(labels, dataArray, dataNames, colorPalette, colorPalette, chartTitle, xLabel, yLabel));

        var ctx = document.getElementById('canvas-main-chart-stacked-bar').getContext('2d');
        window.myLine = new Chart(ctx, createStackedBarChart(labels, dataArray, dataNames, colorPalette, colorPaletteShallow, chartTitle, xLabel, yLabel));

        var ctx = document.getElementById('canvas-main-chart-radar').getContext('2d');
        window.myLine = new Chart(ctx, createRadarChart(labels, dataArray, dataNames, colorPalette, colorPaletteShallow, capacityLabel, xLabel, yLabel));

        var ctx = document.getElementById('canvas-main-chart-doughnut').getContext('2d');
        window.myLine = new Chart(ctx, createDoughnutChart(totalCapacity, dataNames, colorPaletteShallow, colorPaletteShallow, capacityLabel, xLabel, yLabel));

        var ctx = document.getElementById('canvas-main-chart-pie').getContext('2d');
        window.myLine = new Chart(ctx, createPieChart(labels, dataArray, dataNames, colorPaletteShallow, colorPaletteShallow, capacityLabel, xLabel, yLabel));

        var ctx = document.getElementById('canvas-main-chart-polar').getContext('2d');
        window.myLine = new Chart.PolarArea(ctx, createPolarChart(labels, dataArray, dataNames, colorPaletteShallow, colorPaletteShallow, capacityLabel, xLabel, yLabel));

        var ctx = document.getElementById('canvas-main-chart-vertical-bar').getContext('2d');
        window.myLine = new Chart(ctx, createVerticalBarChart(labels, dataArray, dataNames, colorPalette, colorPaletteShallow, chartTitle, xLabel, yLabel));

        var ctx = document.getElementById('canvas-main-chart-horizontal-bar').getContext('2d');
        window.myLine = new Chart(ctx, createHorizontalBarChart(labels, dataArray, dataNames, colorPalette, colorPaletteShallow, chartTitle, xLabel, yLabel));
    };
    </script>
    <br>
    <br>
"@

$dataFragment = ConvertTo-Html -Head $head -Body $database -PostContent $postContent
$dataFragment | Out-String | Out-File $args[1]