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
import { UploadOutlined, InboxOutlined } from "@ant-design/icons";
import {
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

const { Header, Content } = Layout;
const { Title } = Typography;
const { Dragger } = Upload;

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

const App: React.FC = () => {
  const [config, setConfig] = useState<ConfigType>({});
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
                    setConfig={setConfig}
                    selectConfig={(colsKey, colsValue) => {
                      const config = getMatchConfig(colsKey, colsValue);
                      setConfig(config);
                      console.log(config);
                    }}
                  />
                </Card>
              </Col>
              <Col span={12}>
                {Object.keys(config).length === 0 ? null : (
                  <Card title="选择要被修改的文件" bordered={false}>
                    <PickerPanel
                      uploaderButtonText="upload"
                      selectsPlaceHolder={[
                        "选择需要匹配的列",
                        "选择匹配后要填充的列",
                      ]}
                      multiple
                      config={config}
                    />
                  </Card>
                )}
              </Col>
            </Row>
          </Content>
        </Layout>
      </Layout>
    </Space>
  );
};

export default App;
type PickPanelProps = {
  uploaderButtonText: string;
  selectsPlaceHolder: string[];
  multiple?: boolean;
  selectConfig?: (...args: [][]) => void;
  config?: ConfigType;
  setConfig?: React.Dispatch<React.SetStateAction<ConfigType>>;
};
const PickerPanel: FC<PickPanelProps> = ({
  uploaderButtonText,
  selectsPlaceHolder,
  selectConfig,
  config,
  multiple,
  setConfig,
}) => {
  const [sheetOptions, setSheetOptions] = useState<
    { label: string; value: number }[]
  >([]);
  const [sheetIndex, setSheetIndex] = useState<string | number>("");
  const workBookListRef = useRef<
    { workbook: Excel.Workbook; uid: string; type: string; name: string }[]
  >([]);

  const uploadProps: UploadProps<File> = {
    accept:
      "application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    beforeUpload: () => {
      return false;
    },
    async onChange(info) {
      console.log(workBookListRef.current);
      const file = info.file;
      if (file == null) return;
      if (file.status === "removed") {
        workBookListRef.current = workBookListRef.current.filter((it) => {
          return it.uid !== file.uid;
        });
        if (workBookListRef.current.length) return;
        setSheetIndex("");
        setColKey("");
        setColV("");
        setColsHeader([]);
        setSheetOptions([]);
        setConfig?.({});
        setIsOk(false);
        return;
      }
      const workbook = await loadFile(file as unknown as File);
      workBookListRef.current.push({
        workbook,
        uid: file.uid,
        name: file.name,
        type:
          file.type ||
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const configSheets = getSheets(workbook);
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
  const currentSheet = useMemo(() => {
    const { workbook } = workBookListRef.current[0] || {};
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
    if (currentSheet == null) return;
    const title = getRow(currentSheet, 1);
    setColsHeader(title);
  }, [currentSheet]);

  const [isOk, setIsOk] = useState(false);

  useEffect(() => {
    if (currentSheet == null) return;
    if (!colV) return;
    if (!colK) return;
    const colsValue = getColumn(currentSheet, colV);
    const colsKey = getColumn(currentSheet, colK);
    selectConfig?.(colsKey, colsValue);
    if (config == null) return;
    if (!workBookListRef.current.length) return;
    for (const it of workBookListRef.current) {
      const sheet = getSheets(it.workbook)[sheetIndex as number];
      fillValues(sheet, config, colsKey, colV);
    }
    setIsOk(true);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentSheet, colV, colK]);

  const width = 240;

  return (
    <>
      {multiple ? (
        <Dragger {...uploadProps} multiple={multiple}>
          <p className="ant-upload-drag-icon">
            <InboxOutlined />
          </p>
          <p className="ant-upload-text">
            Click or drag file to this area to upload
          </p>
          <p className="ant-upload-hint">
            支持上传多个模板相同的文件，批量处理！
          </p>
        </Dragger>
      ) : (
        <Upload {...uploadProps}>
          <Button icon={<UploadOutlined />}>{uploaderButtonText}</Button>
        </Upload>
      )}
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
      {isOk ? (
        <Button
          type="primary"
          onClick={() => {
            if (!workBookListRef.current.length) return;
            for (const { workbook, type, name } of workBookListRef.current) {
              saveFile(workbook, type, "ok-" + name);
            }
          }}
        >
          导出文件
        </Button>
      ) : null}
    </>
  );
};
