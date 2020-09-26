########################################################################################################

function Get-Request {
    process {
        curl -Uri $PSItem -ContentType 'application/json' | select -ExpandProperty Content | ConvertFrom-Json
    }
}

function Get-Coin {
    param($id = 'akropolis')
    $uri = 'https://api.coingecko.com/api/v3/coins/{0}?localization=false&tickers=false&market_data=false&community_data=false&developer_data=false&sparkline=false' -f $id
    $uri | Get-Request | select id, symbol, name, asset_platform_id, contract_address
}

function Get-Price {
    param($id = 'akropolis', $cur = 'eth', $contract = $null)

    $coin = Get-Coin -id $id

    if([boolean]$coin.contract_address -and $coin.asset_platform_id -eq 'ethereum') {
        $req = 'https://api.coingecko.com/api/v3/simple/token_price/ethereum?contract_addresses={0}&vs_currencies={1}' -f $coin.contract_address, $cur | Get-Request
        [decimal]$req.($coin.contract_address).$cur
    } else {
        $req = 'https://api.coingecko.com/api/v3/simple/price?ids={0}&vs_currencies={1}' -f $id, $cur | Get-Request
        [decimal]$req.$id.$cur
    }
}

function Get-UniPrice {

    param($id = '0x8cb77ea869def8f7fdeab9e4da6cf02897bbf076')

$query = @"
{
    pair(id: "$id"){
        token0 {
            id
            symbol
            name
            derivedETH
        }
        token1 {
            id
            symbol
            name
            derivedETH
        }
        reserve0
        reserve1
        reserveUSD
        trackedReserveETH
        token0Price
        token1Price
        volumeUSD
        txCount
    }
}
"@

    $body = @{
        query     = $query
        #variables = $Variables
    }

    $header = @{
      'content-type' = 'application/json'
      'Accept' = 'application/json'
      #"Authorization" = 'bearer tokenCode'
    }

    $uri = 'https://api.thegraph.com/subgraphs/name/uniswap/uniswap-v2'
    $data = Invoke-RestMethod -Method Post -Uri $uri -Body ($body | ConvertTo-Json) -Headers $header

    if($data.data.pair.token0.symbol -eq 'WETH') {
        #'[{0}] {1} {2}' -f $data.data.pair.token1.symbol, $data.data.pair.token1.derivedETH, $data.data.pair.token0.symbol
        [decimal]$data.data.pair.token1.derivedETH
    } else {
        #'[{0}] {1} {2}' -f $data.data.pair.token0.symbol, $data.data.pair.token0.derivedETH, $data.data.pair.token1.symbol
        [decimal]$data.data.pair.token0.derivedETH
    }
}

########################################################################################################

$shizzlebagUNI = [ordered] @{

    AKRO  = @{ contract='0x8ab7404063ec4dbcfd4598215992dc3f8ec853d7'; pool='0x8cb77ea869def8f7fdeab9e4da6cf02897bbf076' }
    ADEL  = @{ contract='0x94d863173ee77439e4292284ff13fad54b3ba182'; pool='0xad19b7e2b1dac6cea46b18eec731981c08e6f08e' }
    PLU   = @{ contract='0xd8912c10681d8b21fd3742244f44658dba12264e'; pool='0x87c9524237a19338be7dbcac01d6d208ff31136f' }
    DOS   = @{ contract='0x0a913bead80f321e7ac35285ee10d9d922659cb7'; pool='0xdadf443c086f9d3c556ebc57c398a852f6a02898' }
    RING  = @{ contract='0x9469d013805bffb7d3debe5e7839237e535ec483'; pool='0xa32523371390b0cc4e11f6bb236ecf4c2cdea101' }
    UNC   = @{ contract='0xf29e46887ffae92f1ff87dfe39713875da541373'; pool='0x5e64cd6f84d0ee2ad2a84cadc464184e36274e0c' }
    LEND  = @{ contract='0x80fb784b7ed66730e8b1dbd9820afd29931aab03'; pool='0xab3f9bf1d81ddb224a2014e98b238638824bcf20' }
    SNX   = @{ contract='0xc011a73ee8576fb46f5e1c5751ca3b9fe0af2a6f'; pool='0x43ae24960e5534731fc831386c07755a2dc33d47' }
    SWAP  = @{ contract='0xcc4304a31d09258b0029ea7fe63d032f52e44efe'; pool='0xd90a1ba0cbaaaabfdc6c814cdf1611306a26e1f8' }
    MET   = @{ contract='0xa3d58c4e56fedcae3a7c43a725aee9a71f0ece4e'; pool='0x350db0440f1ff18e900dcc8c01180aa00e72cc91' }
    XOR   = @{ contract='0x40fd72257597aa14c7231a7b1aaa29fce868f677'; pool='0x01962144d41415cca072900fe87bbe2992a99f10' }
    PAR   = @{ contract='0x1beef31946fbbb40b877a72e4ae04a8d1a5cee06'; pool='0x6aeebc2f5c979fd5c4361c2d288e55ac6b7e39bb' }
    ANKR  = @{ contract='0x8290333cef9e6d528dd5618fb97a76f268f3edd4'; pool='0x5201883feeb05822ce25c9af8ab41fc78ca73fa9' }
    VIDT  = @{ contract='0x445f51299ef3307dbd75036dd896565f5b4bf7a5'; pool='0x355df254ecd2dfb26c02b7331c2bbe09259b498d' }
    PUX   = @{ contract='0xe277ac35f9d327a670c1a3f3eec80a83022431e4'; pool='0x48631f2eef62f1ace0600e5db38cfbf77f64a3e8' }
    TEND  = @{ contract='0x1453dbb8a29551ade11d89825ca812e05317eaeb'; pool='0xcfb8cf118b4f0abb2e8ce6dbeb90d6bc0a62693d' }
    OM    = @{ contract='0x2baecdf43734f22fd5c152db08e3c27233f0c7d2'; pool='0x99b1db3318aa3040f336fb65c55400e164ddcd7f' }
    XFT   = @{ contract='0xabe580e7ee158da464b51ee1a83ac0289622e6be'; pool='0x2b9e92a5b6e69db9fedc47a4c656c9395e8a26d2' }
    DEXT  = @{ contract='0x26ce25148832c04f3d7f26f32478a9fe55197166'; pool='0x37a0464f8f4c207b54821f3c799afd3d262aa944' }
    XRT   = @{ contract='0x7de91b204c1c737bcee6f000aaa6569cf7061cb7'; pool='0x3185626c14acb9531d19560decb9d3e5e80681b1' }
    CHSB  = @{ contract='0xba9d4199fab4f26efe3551d490e3821486f135ba'; pool='0x8d9b9e25b208cac58415d915898c2ffa3a530aa1' }
    UNI   = @{ contract='0x1f9840a85d5af5bf1d1762f925bdaddc4201f984'; pool='0xd3d2e2692501a5c9ca623199d38826e513033a17' }
    CHART = @{ contract='0x1d37986f252d0e349522ea6c3b98cb935495e63e'; pool='0x960d228bb345fe116ba4cba4761aab24a5fa7213' }
}

$shizzlebag = [ordered] @{

    BTC   = @{ cell = 'G31'; cur = 'eur'; id = 'bitcoin' }
    ETH   = @{ cell = 'R2';  cur = 'eur'; id = 'ethereum' }
    AKRO  = @{ cell = 'O2';  cur = 'eth'; id = 'akropolis' }
    ADEL  = @{ cell = 'O2';  cur = 'eth'; id = 'akropolis-delphi' }
    PLU   = @{ cell = 'O2';  cur = 'eth'; id = 'pluton' }
    DOS   = @{ cell = 'O2';  cur = 'eth'; id = 'dos-network' }
    RING  = @{ cell = 'O2';  cur = 'eth'; id = 'darwinia-network-native-token' }
    UNC   = @{ cell = 'O17'; cur = 'eth'; id = 'unicrypt' }
    LEND  = @{ cell = 'O15'; cur = 'eth'; id = 'ethlend' }
    SNX   = @{ cell = 'O15'; cur = 'eth'; id = 'havven' }
    SWAP  = @{ cell = 'O15'; cur = 'eth'; id = 'trustswap' }
    MET   = @{ cell = 'M15'; cur = 'eth'; id = 'metronome' }
    XOR   = @{ cell = 'O15'; cur = 'eth'; id = 'sora' }
    PAR   = @{ cell = 'O15'; cur = 'eth'; id = 'parachute' }
    ANKR  = @{ cell = 'O15'; cur = 'eth'; id = 'ankr' }
    VIDT  = @{ cell = 'O15'; cur = 'eth'; id = 'v-id-blockchain' }
    PUX   = @{ cell = 'M15'; cur = 'eth'; id = 'polypux' }
    TEND  = @{ cell = 'M15'; cur = 'eth'; id = 'tendies' }
    OM    = @{ cell = 'O2';  cur = 'eth'; id = 'mantra-dao' }
    XFT   = @{ cell = 'O15'; cur = 'eth'; id = 'offshift' }
    DEXT  = @{ cell = 'O15'; cur = 'eth'; id = 'idextools' }
    XRT   = @{ cell = 'O15'; cur = 'eth'; id = 'robonomics-network' }
    CHSB  = @{ cell = 'O15'; cur = 'eth'; id = 'swissborg' }
    UNI   = @{ cell = 'O2';  cur = 'eth'; id = 'uniswap' }
    CHART = @{ cell = 'O2';  cur = 'eth'; id = 'chartex' }
}

########################################################################################################

$file = '/Users/schyvensjim/Documents/test.xlsx'
$xlsx = Open-ExcelPackage -Path $file

foreach($key in $shizzlebag.Keys) {

    $coin = $shizzlebag[$key]

    if($key -in $shizzlebagUNI.Keys) {
        $price = Get-UniPrice -id $shizzlebagUNI[$key].pool
    } else {
        $price = Get-Price -id $coin.id -cur $coin.cur
    }

    $sheet = $xlsx.Workbook.Worksheets["Sheet1"]
    $val = [decimal]$sheet.Cells[$coin.cell].Value
    $color = 'DarkGreen'

    if($val -gt $price) {
        $color = 'DarkRed'
    }

    $diff = $price - $val
    $change = '0'
    if([boolean]$val) {
        $change = ($diff / $val) * 100
    }

    "[{0}]`t{1}`t{2:n2}%" -f $key, $diff, $change | Write-Host -BackgroundColor $color

    $sheet.Cells[$coin.cell].Value = $price
}

Close-ExcelPackage -ExcelPackage $xlsx -Show

########################################################################################################
