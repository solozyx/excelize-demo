package main

import (
	"fmt"
	"strconv"
	"sync"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/sirupsen/logrus"
	"gopkg.in/mgo.v2"
	"gopkg.in/mgo.v2/bson"
)

func init() {
	// log
	logrus.SetLevel(logrus.DebugLevel)
	logrus.SetFormatter(&logrus.JSONFormatter{})
	// mgo
	getMgoSession()
}

func main() {
	// csvImportLarge()
	// csvImportSmall()
	// csvExportAppend()

	// xlsxImport()
	// xlsxExportSample()
	// xlsxExportChart()
	// xlsxExportPicture()

	exportType := "all"
	orders, err := GetOrders(ExportOrdersType(exportType), 2)
	if err != nil {

	}
	if len(orders) == 0 {
		return
	}
	logrus.Debugf("product order [all] list count=%d", len(orders))

	_, _ = xlsxExportProductOrders(orders)
}

func xlsxExportProductOrders(orders []MgoProductOrder) (int, error) {
	var err error
	realCount := 0

	// 表头
	categories := map[string]string{
		"A1": "订单号",
		"B1": "用户id",
		"C1": "商品名称",
		"D1": "总金额(元)",
		"E1": "商品现价(元)",
		"F1": "商品原价(元)",
		"G1": "运费(元)",
		"H1": "购买数量",
		"I1": "支付状态",
		"J1": "付款时间",
		"K1": "收货人",
		"L1": "手机",
		"M1": "地区",
		"N1": "详细地址",
		"O1": "发货状态",
		"P1": "操作员",
		"Q1": "快递商",
		"R1": "快递单号",
	}
	f := excelize.NewFile()
	for k, v := range categories {
		err = f.SetCellValue("Sheet1", k, v)
		if err != nil {
			return 0, err
		}
	}

	// 数据
	for i := 0; i < len(orders); i++ {
		values := map[string]interface{}{
			fmt.Sprintf("A%d", i+2): orders[i].OrderId,
			fmt.Sprintf("B%d", i+2): orders[i].UserId,
			fmt.Sprintf("C%d", i+2): orders[i].ProductName,
			fmt.Sprintf("D%d", i+2): fen2yuan(orders[i].FeeCents),
			fmt.Sprintf("E%d", i+2): fen2yuan(orders[i].PresentPrice),
			fmt.Sprintf("F%d", i+2): fen2yuan(orders[i].OriginalPrice),
			fmt.Sprintf("G%d", i+2): fen2yuan(orders[i].Freight),
			fmt.Sprintf("H%d", i+2): orders[i].Quantity,
			fmt.Sprintf("I%d", i+2): orders[i].Status.String(),
			fmt.Sprintf("J%d", i+2): formatPayAtStr(orders[i].PayAt),
			fmt.Sprintf("K%d", i+2): orders[i].Name,
			fmt.Sprintf("L%d", i+2): orders[i].Phone,
			fmt.Sprintf("M%d", i+2): orders[i].Area,
			fmt.Sprintf("N%d", i+2): orders[i].Address,
			fmt.Sprintf("O%d", i+2): orders[i].SendStatus,
			fmt.Sprintf("P%d", i+2): orders[i].Operator,
			fmt.Sprintf("Q%d", i+2): orders[i].Express,
			fmt.Sprintf("R%d", i+2): orders[i].ExpressNo,
		}
		for k, v := range values {
			err = f.SetCellValue("Sheet1", k, v)
			if err != nil {
				return 0, err
			}
		}

		// 实际导出条数
		realCount++
	}

	// "2020-04-09T18:16:07+08:00" -> "20200409_18_16_07"
	formatStr := "20060102_15_04_05"
	formatTime, _ := time.Parse(time.RFC3339, time.Now().Format(time.RFC3339))
	timeStr := formatTime.Format(formatStr)
	// fileName := fmt.Sprintf("%sProductOrder_%s.xlsx", "/opt/", timeStr)
	fileName := fmt.Sprintf("%sProductOrder_%s.xlsx", "./product_orders/", timeStr)

	// 保存Excel文件
	err = f.SaveAs(fileName)
	if err != nil {
		return 0, err
	}

	return realCount, nil
}

func fen2yuan(fen int) string {
	fenStr := strconv.Itoa(fen)
	if len(fenStr) == 1 {
		return fmt.Sprintf("0.0%s", fenStr)
	} else if len(fenStr) == 2 {
		return fmt.Sprintf("0.%s", fenStr)
	}
	firstStr := fenStr[:len(fenStr)-2]
	lastStr := fenStr[len(fenStr)-2:]
	yuan := fmt.Sprintf("%s.%s", firstStr, lastStr)
	return yuan
}

func formatPayAtStr(str string) string {
	// "2020-04-09T18:16:07+08:00" -> "2020.04.09 18:16:07"
	formatStr := "2006-01-02 15:04:05"
	timeStr, _ := time.Parse(time.RFC3339, str)
	s := timeStr.Format(formatStr)
	if s == "0001-01-01 00:00:00" {
		return ""
	}
	return s
}

const (
	_mgoProductOrder = "mgo_product_order"
)

// 根据发货状态查询订单列表
func GetOrders(exportType ExportOrdersType, count int) ([]MgoProductOrder, error) {
	session := mgoSess.Copy()
	defer session.Close()

	collection := session.DB(_mgoDB).C(_mgoProductOrder)
	list := make([]MgoProductOrder, 0)
	var err error

	if exportType == ExportOrdersTypeAll {
		err = collection.Find(bson.M{"status": OrderStatusSucceeded}).Limit(count).All(&list)
		if err != nil {
			logrus.Errorf("get orders [all] list error=%v", err.Error())
			return nil, err
		}
	} else if exportType == ExportOrdersTypeUnSend {
		err = collection.Find(bson.M{"status": OrderStatusSucceeded, "send_status": SendStatusWaiting}).Limit(count).All(&list)
		if err != nil {
			logrus.Errorf("get orders [SendStatusWaiting] list error=%v", err.Error())
			return nil, err
		}
	} else if exportType == ExportOrdersTypeSent {
		err = collection.Find(bson.M{"status": OrderStatusSucceeded, "send_status": SendStatusSucceeded}).Limit(count).All(&list)
		if err != nil {
			logrus.Errorf("get orders [SendStatusSucceeded] list error=%v", err.Error())
			return nil, err
		}
	}

	return list, nil
}

type MgoProductOrder struct {
	ID            bson.ObjectId `bson:"_id" json:"id"`
	OrderId       string        `bson:"order_id" json:"order_id"`               // 订单编号 SN1234567890
	OrderType     OrderType     `bson:"order_type" json:"order_type"`           // 订单类型
	Payment       string        `bson:"payment" json:"payment"`                 // 支付方式
	UserId        int           `bson:"user_id" json:"user_id"`                 // 用户ID
	PrepayId      string        `bson:"prepay_id" json:"prepay_id"`             // 预支付交易会话标识
	TransactionId string        `bson:"transaction_id" json:"transaction_id"`   // 微信支付订单号
	ReceiptData   string        `bson:"receipt_data" json:"receipt_data"`       // ApplePay支付凭证
	FeeCents      int           `bson:"fee_cents" json:"fee_cents"`             // 微信支付金额,单位分
	ProductId     int           `bson:"product_id" json:"product_id"`           // 商品ID
	ProductName   string        `bson:"product_name" json:"product_name"`       // 商品名称
	Sku           string        `bson:"sku" json:"sku"`                         // 商品名称
	ThumbImageUrl string        `bson:"thumb_image_url" json:"thumb_image_url"` // 商品缩略图
	PresentPrice  int           `bson:"present_price" json:"present_price"`     // 商品现价,单位分
	OriginalPrice int           `bson:"original_price" json:"original_price"`   // 商品原价,单位分
	Quantity      int           `bson:"quantity" json:"quantity"`               // 商品购买数量
	Freight       int           `bson:"freight" json:"freight"`                 // 订单运费
	Status        OrderStatus   `bson:"status" json:"status"`                   // 订单状态
	Name          string        `bson:"name" json:"name"`                       // 收货人姓名
	Phone         string        `bson:"phone" json:"phone"`                     // 收货人电话
	Area          string        `bson:"area" json:"area"`                       // 收货人地区
	Address       string        `bson:"address" json:"address"`                 // 收货人详细地址
	LotteryTimes  int           `bson:"lottery_times" json:"lottery_times"`     // 该订单增加的抽奖次数
	LotteryStatus LotteryStatus `bson:"lottery_status" json:"lottery_status"`   // 抽奖次数增加状态
	SendStatus    SendStatus    `bson:"send_status" json:"send_status"`         // 发货状态
	Operator      string        `bson:"operator" json:"operator"`               // 操作员
	Express       string        `bson:"express" json:"express"`                 // 快递商
	ExpressNo     string        `bson:"express_no" json:"express_no"`           // 快递单号
	CreateAt      string        `bson:"create_at" json:"create_at"`             // 创建时间 e.g. 2006-01-02T15:04:05Z07:00
	PayAt         string        `bson:"pay_at" json:"pay_at"`                   // 完成支付时间 e.g. 2006-01-02T15:04:05Z07:00
	UpdateAt      string        `bson:"update_at" json:"update_at"`             // 更新时间 e.g. 2006-01-02T15:04:05Z07:00
}

type LotteryStatus uint

const (
	LotteryStatusWaiting LotteryStatus = iota
	LotteryStatusSucceeded
	LotteryStatusFailed
)

type SendStatus uint

const (
	SendStatusWaiting SendStatus = iota
	SendStatusSucceeded
)

type OrderType uint

const (
	OrderTypeCrowdfunding OrderType = iota // 购买打赏
	OrderTypeVip
	OrderTypeGift      // 购买礼物
	OrderTypeAIBalance // 充值AI
	OrderTypeIdol      // 购买模型
	OrderTypeProduct   // 购买商品
)

func (p OrderType) String() string {
	return [...]string{"明星众筹", "会员订阅", "赠送礼物", "充值AI", "模型购买", "实物商品"}[p]
}

type OrderStatus uint

const (
	OrderStatusWaiting OrderStatus = iota
	OrderStatusSucceeded
	OrderStatusFailed
)

func (p OrderStatus) String() string {
	return [...]string{"未支付", "已支付", "支付失败"}[p]
}

func (p SendStatus) String() string {
	return [...]string{"未发货", "已发货"}[p]
}

type ExportOrdersType string

const (
	ExportOrdersTypeAll    ExportOrdersType = "all"
	ExportOrdersTypeUnSend ExportOrdersType = "unsend"
	ExportOrdersTypeSent   ExportOrdersType = "sent"
)

var (
	mgoSess *mgo.Session
	rwLock  sync.RWMutex
)

const (
	_mgoHost = "127.0.0.1:27019"
	_mgoDB   = "MzDB"
)

func getMgoSession() *mgo.Session {
	if mgoSess == nil {
		rwLock.Lock()
		defer rwLock.Unlock()
		if mgoSess == nil {
			mgoSess = connectMongo()
		}
	}
	return mgoSess
}

func connectMongo() *mgo.Session {
	s, err := mgo.DialWithInfo(&mgo.DialInfo{
		Addrs:    []string{_mgoHost},
		Database: _mgoDB,
		Username: "",
		Password: "",
	})
	if err != nil {
		panic(err)
	}
	return s
}
