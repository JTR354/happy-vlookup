import React, { FC, useEffect, useMemo, useRef, useState } from "react";
import {
  Layout,
  Space,
  Typography,
  Select,
  Button,
  Upload,
  Card,
  Row,
  Col,
} from "antd";
import type { UploadProps } from "antd/es/upload/interface";
import { UploadOutlined } from "@ant-design/icons";
import {
  ColsType,
  ConfigType,
  fillValues,
  getColumn,
  getMatchConfig,
  getRow,
  getSheets,
  loadFile,
  saveFile,
} from "./hooks";
import type Excel from "exceljs";

const { Header, Footer, Content } = Layout;
const { Title } = Typography;

const headerStyle: React.CSSProperties = {
  textAlign: "center",
  // color: "#fff",
  height: "auto",
  paddingInline: 50,
  lineHeight: "64px",
  backgroundColor: "#7dbcea",
};

const siderStyle: React.CSSProperties = {
  color: "#fff",
  backgroundColor: "#3ba0e9",
};

const footerStyle: React.CSSProperties = {
  textAlign: "center",
  color: "#fff",
  backgroundColor: "#7dbcea",
};

const App: React.FC = () => {
  const [config, setConfig] = useState<ConfigType>({});
  const workbookRef = useRef<{ workbook?: Excel.Workbook; fileType?: string }>(
    {}
  );
  return (
    <Space direction="vertical" style={{ width: "100%" }} size={[0, 48]}>
      <Layout>
        <Header style={headerStyle}>
          <Title>Happy vlookup</Title>
        </Header>
        <Layout>
          <Content style={siderStyle}>
            <Row gutter={16}>
              <Col span={12}>
                <Card title="选择模板文件" bordered={false}>
                  <PickerPanel
                    uploaderButtonText="upload"
                    selectsPlaceHolder={[
                      "选择需要匹配的列",
                      "选择匹配后要填充的列",
                    ]}
                    selectConfig={(colsKey, colsValue) => {
                      const config = getMatchConfig(colsKey, colsValue);
                      setConfig(config);
                    }}
                  />
                </Card>
              </Col>
              <Col span={12}>
                <Card title="选择要被修改的文件" bordered={false}>
                  <PickerPanel
                    uploaderButtonText="upload"
                    selectsPlaceHolder={[
                      "选择需要匹配的列",
                      "选择匹配后要填充的列",
                    ]}
                    // selectConfig={(colsKey, colsValue) => {
                    //   fillValues();
                    // }}
                    selectEnd={(
                      workbook,
                      index,
                      colsKey,
                      fillCol,
                      fileType
                    ) => {
                      if (Object.keys(config).length === 0) {
                        alert("请先选择模板");
                        return;
                      }
                      workbookRef.current = { workbook, fileType };
                      const sheet = getSheets(workbook)[index];
                      fillValues(sheet, config, colsKey, fillCol);
                    }}
                  />
                </Card>
              </Col>
            </Row>
          </Content>
        </Layout>
        <Footer style={footerStyle}>
          <Button
            type="primary"
            onClick={() => {
              const { workbook, fileType = "" } = workbookRef.current;
              if (workbook == null) return;
              saveFile(workbook, fileType);
            }}
          >
            确定，并现在文件
          </Button>
        </Footer>
      </Layout>
    </Space>
  );
};

export default App;
type PickPanelProps = {
  uploaderButtonText: string;
  selectsPlaceHolder: string[];
  selectConfig?: (...args: [][]) => void;
  selectEnd?: (
    sheet: Excel.Workbook,
    configIndex: number,
    colsKey: ColsType,
    fillCol: string,
    fileType: string
  ) => void;
};
const PickerPanel: FC<PickPanelProps> = ({
  uploaderButtonText,
  selectsPlaceHolder,
  selectConfig,
  selectEnd,
}) => {
  const [sheetOptions, setSheetOptions] = useState<
    { label: string; value: number }[]
  >([]);
  const fileType = useRef<string | unknown>("");
  const [sheetIndex, setSheetIndex] = useState<string | number>("");
  const currentWorkBookRef = useRef<Excel.Workbook | null>(null);
  const uploadProps: UploadProps<File> = {
    beforeUpload: () => {
      return false;
    },
    async onChange(info) {
      const file = info.file;
      fileType.current = file.type;
      if (file == null) return;
      if (file.status === "removed") {
        setSheetIndex("");
        setColKey("");
        setColV("");
        setColsHeader([]);
        return;
      }
      const workbook = await loadFile(file as unknown as File);
      const configSheets = getSheets(workbook);
      currentWorkBookRef.current = workbook;
      setSheetOptions(
        configSheets.map((it, index) => {
          return {
            label: it.name,
            value: index,
          };
        })
      );
      setSheetIndex(0);
    },
  };
  const sheet = useMemo(() => {
    const workbook = currentWorkBookRef.current;
    if (workbook == null) return;
    if (typeof sheetIndex === "string") return;
    return getSheets(workbook as Excel.Workbook)[+sheetIndex || 0];
  }, [sheetIndex]);
  const [colsHeader, setColsHeader] = useState<
    { label: string; value: string }[]
  >([]);
  const [colK, setColKey] = useState("");
  const [colV, setColV] = useState("");
  useEffect(() => {
    if (sheet == null) return;
    const title = getRow(sheet, 1);
    setColsHeader(title);
  }, [sheet]);

  useEffect(() => {
    if (sheet == null) return;
    if (!colV) return;
    if (!colK) return;
    const colsValue = getColumn(sheet, colV);
    const colsKey = getColumn(sheet, colK);
    if (currentWorkBookRef.current == null) return;
    selectConfig?.(colsKey, colsValue);
    selectEnd?.(
      currentWorkBookRef.current as unknown as Excel.Workbook,
      sheetIndex as number,
      colsKey,
      colV,
      fileType.current as string
    );
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sheet, colV, colK]);

  const width = 240;

  return (
    <>
      <Upload {...uploadProps}>
        <Button icon={<UploadOutlined />}>{uploaderButtonText}</Button>
      </Upload>
      <div>
        <label>
          <span>选择工作表：</span>
          <Select
            style={{ width }}
            onChange={(i) => {
              setSheetIndex(sheetOptions[+i]?.value || 0);
            }}
            value={sheetIndex}
            options={sheetOptions}
            placeholder="选择工作表"
          />
        </label>
      </div>
      <div>
        <label>
          <span>{selectsPlaceHolder[0]} :</span>
          <Select
            style={{ width }}
            value={colK}
            onChange={(value) => {
              setColKey(value);
            }}
            options={colsHeader}
            placeholder={selectsPlaceHolder[0]}
          />
        </label>
      </div>
      <div>
        <label>
          <span>{selectsPlaceHolder[1]}: </span>
          <Select
            style={{ width }}
            value={colV}
            onChange={(value) => {
              setColV(value);
            }}
            options={colsHeader}
            placeholder={selectsPlaceHolder[1]}
          />
        </label>
      </div>
    </>
  );
};
