package com.rni.serviceImpl;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.sql.Timestamp;
import java.time.Instant;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Optional;
import java.util.Set;

import javax.validation.Valid;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.PageRequest;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.security.core.userdetails.UserDetails;
import org.springframework.stereotype.Service;

import com.rni.models.APARDetailsEO;
import com.rni.models.DayToDayEmployeeTransferEO;
import com.rni.models.DayToDayHeadMasterEO;
import com.rni.models.EmployeeManagementEO;
import com.rni.models.ExpenditureEo;
import com.rni.models.GrantMasterEO;
import com.rni.models.GrantReceivedEo;
import com.rni.models.ITEMRequestFormEO;
import com.rni.models.ImprestRequestEO;
import com.rni.models.IndividualItemRecievedEO;
import com.rni.models.ItemGroupMasterEO;
import com.rni.models.ItemTypeMasterEO;
import com.rni.models.StockRegisterEO;
import com.rni.models.StoreItemMasterEO;
import com.rni.models.StoreItemReceivedEO;
import com.rni.models.StoreSectionBillEO;
import com.rni.models.StoreSectionItemRequestEO;
import com.rni.models.StoreStockEntryInEo;
import com.rni.models.UserEO;
import com.rni.models.VenderEo;
import com.rni.repository.AddBillRepository;
import com.rni.repository.AddVenderRepository;
import com.rni.repository.AparDetailsRepository;
import com.rni.repository.DayToDayEmployeeTransferRepository;
import com.rni.repository.DayToDayHeadMasterRepository;
import com.rni.repository.DayToDayOperationRepository;
import com.rni.repository.DayToDayOperationRepositoryImpl;
import com.rni.repository.DayToDayOperationStockRegisterRepository;
import com.rni.repository.EmployeeManagementRepository;
import com.rni.repository.ExpenditureRepository;
import com.rni.repository.GarntMasterListRepository;
import com.rni.repository.GrantMasterRepository;
import com.rni.repository.GrantReceivedEoRepository;
import com.rni.repository.ImprestRequestRepo;
import com.rni.repository.IndividualItemRecievedRepository;
import com.rni.repository.ItemCategoryMasterRepository;
import com.rni.repository.ItemGroupMasterRepository;
import com.rni.repository.ItemMasterEntryFormDataRepo;
import com.rni.repository.ItemMasterEntryFormRepo;
import com.rni.repository.ItemTypeMasterRepository;
import com.rni.repository.StockInRepo;
import com.rni.repository.StoreItemRepository;
import com.rni.repository.StoreSectionItemRequestRepo;
import com.rni.repository.UserRepository;
import com.rni.request.dto.AddBillDto;
import com.rni.request.dto.AddVendorRequestDto;
import com.rni.request.dto.AparDetailsRequestDto;
import com.rni.request.dto.DayToDayEmployeeTransferRequestDto;
import com.rni.request.dto.DayToDayHeadMasterRequestDto;
import com.rni.request.dto.DayToDayImprestReqDto;
import com.rni.request.dto.DayToDaySearchImprestAccCompletedAndRaisedReqDto;
import com.rni.request.dto.DayToDaySearchImprestAccReqDto;
import com.rni.request.dto.EmployeeBasicInfoRequestDto;
import com.rni.request.dto.EmployeeManagementDto;
import com.rni.request.dto.EmployeeManagementRequestDto;
import com.rni.request.dto.EmployeeProfessionalDetailsRequestDto;
import com.rni.request.dto.ExpenditureRequestDto;
import com.rni.request.dto.GenerateItemCodeRequestDto;
import com.rni.request.dto.GrantMasterDto;
import com.rni.request.dto.GrantReceivedListRequestDto;
import com.rni.request.dto.GrantReceivedRequestDto;
import com.rni.request.dto.HeadMasterDetailByIdRequesteDto;
import com.rni.request.dto.ITEMRequestFormDto;
import com.rni.request.dto.ItemGroupMasterRequestDto;
import com.rni.request.dto.ItemIssueFormDetailRequestDto;
import com.rni.request.dto.ItemMasterEntryFormDataDto;
import com.rni.request.dto.ItemMasterEntryFormListRequestDTO;
import com.rni.request.dto.ItemNameRequestDto;
import com.rni.request.dto.ItemReceivedEntryFormDto;
import com.rni.request.dto.ItemRequestDetailsIssueForm;
import com.rni.request.dto.ItemRequestFormForApprovalDto;
import com.rni.request.dto.ItemTypeMasterDTO;
import com.rni.request.dto.ListOfBillFromStoreRequestDto;
import com.rni.request.dto.PaginationRequestDto;
import com.rni.request.dto.PendingAndApprovedItemRequestsDto;
import com.rni.request.dto.SaveExpenditureRequestDto;
import com.rni.request.dto.SearchItemNameRequestDto;
import com.rni.request.dto.StockEntryListFormRequestDto;
import com.rni.request.dto.StockInDto;
import com.rni.request.dto.StockRegisterRequestDto;
import com.rni.request.dto.UpdateExpenditureRequestDto;
import com.rni.request.dto.UpdateHeadMasterRequestDto;
import com.rni.request.dto.UpdateImprestACRequestsFromEmployeeRequestDto;
import com.rni.request.dto.UpdateItemRequestsFromEmployeesRequestDto;
import com.rni.request.dto.UpdateListOfBillsFromStoreRequestDto;
import com.rni.request.dto.UpdatePendingApprovedITEMRequestsDto;
import com.rni.request.dto.VenderGstAndGemNumberRequestDto;
import com.rni.response.dto.ConsumableItemsStockRegisterResponseDto;
import com.rni.response.dto.DayToDayCommonResponseDto;
import com.rni.response.dto.DayToDayEmployeeTransferViewResDto;
import com.rni.response.dto.DayToDayEmployeetransfeResDto;
import com.rni.response.dto.DayToDayHeadMasterResponseDto;
import com.rni.response.dto.DayToDayListOfItemMasterResDto;
import com.rni.response.dto.DayToDaySearchImprestAccResDto;
import com.rni.response.dto.DropdownUUIDDto;
import com.rni.response.dto.EmployeeManagementResponseDto;
import com.rni.response.dto.EmployeeMasterResponseDto;
import com.rni.response.dto.ExpenditureListResponseDto;
import com.rni.response.dto.ExpenditureResponseDto;
import com.rni.response.dto.GenerateItemCodeResponse;
import com.rni.response.dto.GetDataOfImprestRequestFormResponseDto;
import com.rni.response.dto.GrantMasterResponseDto;
import com.rni.response.dto.ItemGroupMasterResponseDto;
import com.rni.response.dto.ItemIssueFormDetailResponseDto;
import com.rni.response.dto.ItemMasterEntryFormResponseDTO;
import com.rni.response.dto.ItemNameResponseDto;
import com.rni.response.dto.ItemRequestsFromEmployeesResponseDto;
import com.rni.response.dto.ItemTypeMasterResponseDTO;
import com.rni.response.dto.ItemTypeResDto;
import com.rni.response.dto.ListOfAsoRaisedAndCompletedResponseDto;
import com.rni.response.dto.ListOfBillsFromStoreResponseDto;
import com.rni.response.dto.ListOfItemRequestASOResDTO;
import com.rni.response.dto.ListOfRaisedAndCompletedResponseDto;
import com.rni.response.dto.NonConsumableItemsStockRegisterResponseDto;
import com.rni.response.dto.PendingAndApproveItemResponseDto;
import com.rni.response.dto.SearchItemNameResponseDto;
import com.rni.response.dto.StockEntryListFormResponseDto;
import com.rni.response.dto.StoreItemRecieveIssueForm;
import com.rni.response.dto.VenderGstAndGemNumberResponseDto;
import com.rni.security.jwt.JwtUtils;
import com.rni.service.DayToDayOperationService;

@Service
public class DayToDayOperationServiceImpl implements DayToDayOperationService {

	@Autowired
	private ItemMasterEntryFormRepo itemMasterEntryFormRepo;

	@Autowired
	private DayToDaySearchItemNameRepo dayToDaySearchItemNameRepo;

	@Autowired
	private DayToDayOperationStockRegisterRepository dayToDayOperationStockRegisterRepository;
	@Autowired
	private IndividualItemRecievedRepository individualItemRecievedRepository;

	@Autowired
	private ItemMasterEntryFormDataRepo itemMasterEntryFormDataRepo;

	@Autowired
	private StoreSectionItemRequestRepo storeSectionItemRequestRepo;

	@Autowired
	private DayToDayOperationRepository dayToDayOperationRepository;

	@Autowired
	private UserRepository userRepository;

	@Autowired
	private ImprestRequestRepo imprestRepo;

	@Autowired
	private DayToDayEmployeeTransferRepository dayTodayEmployeeTransferRepository;

	@Autowired
	private DayToDayHeadMasterRepository dayToDayHeadMasterRepository;

	@Autowired
	private AddVenderRepository addVenderRepository;

	@Autowired
	private DayToDayOperationRepositoryImpl dayToDayOperationRepositoryImpl;

	@Autowired
	private StockInRepo stockInRepo;

	@Autowired
	private AddBillRepository addBillRepository;

	@Autowired
	private ExpenditureRepository expenditureRepository;

	@Autowired
	private JwtUtils jwtUtils;

	@Autowired
	private GrantMasterRepository masterRepository;

	@Autowired
	private EmployeeManagementRepository employeeManagementRepository;

	@Autowired
	private GarntMasterListRepository garntMasterListRepository;

	@Autowired
	private GrantReceivedEoRepository grantReceivedRepository;

	@Autowired
	StoreItemRepository storeItemRepository;

	@Autowired
	ItemTypeMasterRepository itemTypeMasterRepository;

	@Autowired
	private ItemGroupMasterRepository itemGroupMasterRepository;
	@Autowired
	private ItemCategoryMasterRepository itemCategoryMasterRepository;

	@Autowired
	private AparDetailsRepository aparDetailsRepository;

	@Override
	public ItemReceivedEntryFormDto saveItemMasterEntryForm(ItemReceivedEntryFormDto dto) {
		StoreItemReceivedEO itemObject = new StoreItemReceivedEO();
		try {
			itemObject.setItemCode(dto.getItemCode());
			itemObject.setItemId(dto.getItemId());
			itemObject.setItemGroupId(dto.getItemGroup());
			itemObject.setItemCategoryId(dto.getItemCategory());
			itemObject.setProcurementId(dto.getProcurement());
			itemObject.setItemQuantity(dto.getQuantity());
			itemObject.setItemPurchaseDate(dto.getItemPurchaseDate());
			itemObject.setWarrantyExpiryDate(dto.getWarrantyExpiryDate());
			itemObject.setRemarks(dto.getRemarks());
			for (int i = 1; i <= dto.getQuantity(); i++) {
				IndividualItemRecievedEO childStoreObj = new IndividualItemRecievedEO();
				Instant instant = Instant.now();
				Timestamp time = Timestamp.from(instant);
				childStoreObj.setItemId(dto.getItemId());
				childStoreObj.setCreatedAt(time);
				childStoreObj.setUpdatedAt(time);
				StringBuilder serialNumberFormate = new StringBuilder();
				LocalDate date = LocalDate.now();
				DateTimeFormatter formatter = DateTimeFormatter.ofPattern("ddMMYYYY");
				serialNumberFormate.append(formatter.format(date));
				String itemNameInitials = dto.getItemName().substring(0, 3);
				serialNumberFormate.append(itemNameInitials);
				System.out.println("hi : " + serialNumberFormate + "1");
				String priviousSerialNUmber = itemMasterEntryFormDataRepo.toCheckSerialNumber();
				if (priviousSerialNUmber != null) {
					String str = priviousSerialNUmber.substring(11);
					Integer k = Integer.valueOf(str);
					Integer j = k + 1;
					serialNumberFormate.append(j);
					childStoreObj.setSerialNo(serialNumberFormate.toString());
				} else {
					childStoreObj.setSerialNo(serialNumberFormate + "" + 1);
				}
				individualItemRecievedRepository.save(childStoreObj);
			}

			Instant instant = Instant.now();
			Timestamp time = Timestamp.from(instant);
			itemObject.setCreatedOn(time);
			itemMasterEntryFormDataRepo.save(itemObject);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return dto;
	}

	@Override
	public void saveImprestReq(DayToDayImprestReqDto dto) {
		ImprestRequestEO imprest = new ImprestRequestEO();
		Authentication authentication = SecurityContextHolder.getContext().getAuthentication();
		UserDetails userDetails = (UserDetails) authentication.getPrincipal();
		UserEO user = userRepository.findByUsername(userDetails.getUsername());
		try {
			imprest.setRequestDate(dto.getRequestDate());
			imprest.setEmployeeId(dto.getEmployeeId());
			imprest.setPurpose(dto.getPurpose());
			imprest.setAmount(dto.getAmount());
			imprest.setStatus("unpaid");
			Instant instant = Instant.now();
			Timestamp time = Timestamp.from(instant);
			imprest.setCreatedOn(time);
			imprest.setCreatedBy(user.getId());
			imprestRepo.save(imprest);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@Override
	public void saveItemRequestForm(List<ITEMRequestFormDto> listOfitemMasterRequestFormDto) {
		Authentication authentication = SecurityContextHolder.getContext().getAuthentication();
		UserDetails userDetails = (UserDetails) authentication.getPrincipal();
		UserEO user = userRepository.findByUsername(userDetails.getUsername());
		Instant instant = Instant.now();
		Timestamp time = Timestamp.from(instant);
		for (ITEMRequestFormDto list : listOfitemMasterRequestFormDto) {

			ITEMRequestFormEO itemRequestFormEo = new ITEMRequestFormEO();

			// itemRequestFormEo.setItemId(list.getItemNameId());
			itemRequestFormEo.setItemId(list.getItemNameId().getId());
			itemRequestFormEo.setItemTypeId(list.getItemTypeId());
			itemRequestFormEo.setItemCategoryId(list.getItemCategoryId());
			itemRequestFormEo.setItemGroupId(list.getItemGroupId());
			itemRequestFormEo.setRequiredQuantity(list.getQuantity());
			itemRequestFormEo.setRequiredDate(list.getRequiredDate());
			itemRequestFormEo.setRemarks(list.getRemarks());
			itemRequestFormEo.setCreatedBy(user.getId());
			itemRequestFormEo.setRemainingIssueQuantity(list.getQuantity());
			itemRequestFormEo.setIssuedQuantity(0);
			itemRequestFormEo.setCreatedOn(time);
			itemRequestFormEo.setStatus("Pending With Store");
			itemRequestFormEo.setActive(true);
			dayToDayOperationRepository.save(itemRequestFormEo);
		}
	}

	/*
	 * @Override public Page<List<ListOfRaisedAndCompletedResponseDto>>
	 * getListOfRaisedAndCompleedRequests( PaginationRequestDto reqDto) { return
	 * dayToDayOperationRepository
	 * .getListOfRaisedAndCompleedRequests(PageRequest.of(reqDto.getPageIndex(),
	 * reqDto.getPageSize())); }
	 */

	@Override
	public List<ListOfAsoRaisedAndCompletedResponseDto> getListOfAsoRaisedAndCompleedRequests() {
		return storeSectionItemRequestRepo.getListOfAsoRaisedAndCompleedRequests();
	}

	@Override
	public List<DayToDayListOfItemMasterResDto> getListOfItemMaster() {
		return itemMasterEntryFormRepo.getListOfItemMaster();
	}

	@Override
	public void UpdateItemRequestsFromEmployees(
			UpdateItemRequestsFromEmployeesRequestDto updateItemRequestsFromEmployeesRequestDto) {
		Instant instant = Instant.now();
		Timestamp time = Timestamp.from(instant);

		ITEMRequestFormEO itemRequestFormEO = dayToDayOperationRepository
				.findById(updateItemRequestsFromEmployeesRequestDto.getItemRequestId()).get();

		if (itemRequestFormEO.getIssuedQuantity() == itemRequestFormEO.getRequiredQuantity()) {
			return;
		}

		if (updateItemRequestsFromEmployeesRequestDto.getAction().equals("Not Available")) {
			itemRequestFormEO.setStatus("Not Available");
			itemRequestFormEO.setUpdatedOn(time);
			dayToDayOperationRepository.save(itemRequestFormEO);
			return;

		}

		itemRequestFormEO.setIssuedQuantity(
				updateItemRequestsFromEmployeesRequestDto.getIssuedQuantity() + itemRequestFormEO.getIssuedQuantity());

		if (itemRequestFormEO.getRequiredQuantity() == itemRequestFormEO.getIssuedQuantity()) {
			itemRequestFormEO.setStatus("Issued");
		}
		if (itemRequestFormEO.getRequiredQuantity() > itemRequestFormEO.getIssuedQuantity()) {
			itemRequestFormEO.setStatus("Partially Issued");
		}
		itemRequestFormEO.setRemainingIssueQuantity(
				itemRequestFormEO.getRequiredQuantity() - itemRequestFormEO.getIssuedQuantity());

		itemRequestFormEO.setRemarkForStoreAdmin(updateItemRequestsFromEmployeesRequestDto.getRemarkForStoreAdmin());
		itemRequestFormEO.setIssueDate(updateItemRequestsFromEmployeesRequestDto.getIssueDate());
		itemRequestFormEO.setUpdatedOn(time);
		dayToDayOperationRepository.save(itemRequestFormEO);

	}

	@Override
	public void saveAddVendorDetail(AddVendorRequestDto addVendorRequestDto) {
		Instant instant = Instant.now();
		Timestamp time = Timestamp.from(instant);
		VenderEo venderEo = new VenderEo();
		venderEo.setCompanyName(addVendorRequestDto.getNameOfTheCompany());
		venderEo.setVenderName(addVendorRequestDto.getNameOfVendorRepresentee());
		venderEo.setEmailId(addVendorRequestDto.getEmailId());
		venderEo.setMobileNumber(addVendorRequestDto.getMobileNumber());
		venderEo.setPanNumber(addVendorRequestDto.getPanNumber());
		venderEo.setGemId(addVendorRequestDto.getGemid());
		venderEo.setGstNumber(addVendorRequestDto.getGstNumber());
		venderEo.setRegisteredAddress(addVendorRequestDto.getRegisteredAddress());
		venderEo.setStateId(addVendorRequestDto.getStateId());
		venderEo.setDistrictId(addVendorRequestDto.getDistrictId());
		venderEo.setPincode(addVendorRequestDto.getPincode());
		venderEo.setCreatedOn(time);
		venderEo.setCreated_by(jwtUtils.getUserFromJwt().getId());
		addVenderRepository.save(venderEo);
	}

	public void saveItemRequestFormForApproval(List<ItemRequestFormForApprovalDto> itemRequestFormForApprovalDto) {
		Authentication authentication = SecurityContextHolder.getContext().getAuthentication();
		UserDetails userDetails = (UserDetails) authentication.getPrincipal();
		UserEO user = userRepository.findByUsername(userDetails.getUsername());
		Instant instant = Instant.now();
		Date date = Date.from(instant);
		Timestamp time = Timestamp.from(instant);
		for (ItemRequestFormForApprovalDto list : itemRequestFormForApprovalDto) {
			StoreSectionItemRequestEO itemRequestFormEo = new StoreSectionItemRequestEO();
			itemRequestFormEo.setItemId(list.getItemId());
			itemRequestFormEo.setItemGroupId(list.getItemGroupId());
			itemRequestFormEo.setRequiredQuantity(list.getQuantity());
			itemRequestFormEo.setRequiredDate(list.getRequiredDate());
			itemRequestFormEo.setRequestDate(date);
			itemRequestFormEo.setRemarks(list.getRemarks());
			itemRequestFormEo.setOfficerUserId(user.getId());
			itemRequestFormEo.setCreatedBy(user.getId());
			itemRequestFormEo.setCreatedOn(time);
			itemRequestFormEo.setApprovedDate(date);
			itemRequestFormEo.setActive(true);
			itemRequestFormEo.setStatus(list.getAction());
			storeSectionItemRequestRepo.save(itemRequestFormEo);
		}
	}

	@Override
	public Page<PendingAndApproveItemResponseDto> getListOfPendingAndApproveItemRequests(
			PendingAndApprovedItemRequestsDto reqDto) {
		Page<PendingAndApproveItemResponseDto> list = null;
		try {
			list = dayToDayOperationRepositoryImpl.getListOfPendingAndApproveItemRequests(reqDto);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;
	}

	@Override
	public Page<ListOfItemRequestASOResDTO> getlistOfItemRequestASO(PendingAndApprovedItemRequestsDto reqDto) {
		Page<ListOfItemRequestASOResDTO> list = null;
		try {
			list = dayToDayOperationRepositoryImpl.getListOfItemRequestReceivedFromASO(reqDto);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;
	}

	@Override
	public Page<ListOfItemRequestASOResDTO> getListOfItemRequestApprovedFromASO(
			PendingAndApprovedItemRequestsDto reqDto) {
		Page<ListOfItemRequestASOResDTO> list = null;
		try {
			list = dayToDayOperationRepositoryImpl.getListOfItemRequestApprovedFromASO(reqDto);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;
	}

	@Override
	public Page<ItemRequestsFromEmployeesResponseDto> getListOfItemRequestsFromEmployees1(
			PendingAndApprovedItemRequestsDto reqDto) {
		Page<ItemRequestsFromEmployeesResponseDto> list = null;
		try {
			list = dayToDayOperationRepositoryImpl.getListOfItemRequestsFromEmployees(reqDto);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;
	}

	@Override
	public Page<ItemRequestsFromEmployeesResponseDto> getListOfPendingItemRequestsFromEmployees(
			PendingAndApprovedItemRequestsDto reqDto) {
		Page<ItemRequestsFromEmployeesResponseDto> list = null;
		try {
			list = dayToDayOperationRepositoryImpl.getListOfItemRequestsFromEmployees(reqDto);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;
	}

	@Override
	public Page<ItemRequestsFromEmployeesResponseDto> getListOfIssuedItemRequestsFromEmployees(
			PendingAndApprovedItemRequestsDto reqDto) {
		Page<ItemRequestsFromEmployeesResponseDto> list = null;
		try {
			list = dayToDayOperationRepositoryImpl.getListOfIssuedItemRequestsFromEmployees(reqDto);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;
	}

	@Override
	public ByteArrayInputStream generateExcelSheetForPendingAndApprovedItemRequests(int pageNumber, int pageSize,
			String itemName, String itemCategory, Date toDate, Date fromDate) throws IOException {
		PendingAndApprovedItemRequestsDto dto = new PendingAndApprovedItemRequestsDto();
		dto.setFromDate(fromDate);
		dto.setToDate(toDate);
		dto.setItemName(itemName);
		dto.setItemCategory(itemCategory);
		dto.setPageNumber(pageNumber);
		dto.setPageSize(pageSize);
		Page<PendingAndApproveItemResponseDto> page = dayToDayOperationRepositoryImpl
				.getListOfPendingAndApproveItemRequests(dto);
		List<PendingAndApproveItemResponseDto> list = page.getContent();
		XSSFWorkbook workbook = new XSSFWorkbook();
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		Font font = workbook.createFont();
		font.setColor(IndexedColors.YELLOW.getIndex());
		style.setFont(font);
		XSSFSheet sheet = workbook.createSheet("Pending And Approved Item Request");
		Row titleRow = sheet.createRow(0);
		Cell titleCell = titleRow.createCell(0);
		titleCell.setCellValue("Pending And Approved Item Request");
		CellStyle titleCellStyle = workbook.createCellStyle();
		Font titleFont = workbook.createFont();
		titleFont.setBold(true);
		titleFont.setFontName("Arial");
		titleCellStyle.setFont(titleFont);
		titleCellStyle.setBorderBottom(BorderStyle.THIN);
		titleCellStyle.setAlignment(HorizontalAlignment.CENTER);
		titleCell.setCellStyle(titleCellStyle);
		Row header = sheet.createRow(1);
		CellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setBorderBottom(BorderStyle.THIN);
		headerStyle.setAlignment(HorizontalAlignment.CENTER);
		headerStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		XSSFFont font1 = ((XSSFWorkbook) workbook).createFont();
		font.setFontName("Arial");
		font.setFontHeightInPoints((short) 10);
		font1.setBold(true);
		headerStyle.setFont(font1);
		Cell headerCell = header.createCell(0);
		headerCell.setCellValue("Item Name");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(1);
		headerCell.setCellValue("Item Group");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(2);
		headerCell.setCellValue("Required Quantity");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(3);
		headerCell.setCellValue("Required Date");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(4);
		headerCell.setCellValue("Request Date");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(5);
		headerCell.setCellValue("Remarks");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(6);
		headerCell.setCellValue("Action");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(7);
		headerCell.setCellValue("Remarks from Store PR");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(8);
		headerCell.setCellValue("Approved Date");
		headerCell.setCellStyle(headerStyle);

		CellStyle rowStyle = workbook.createCellStyle();
		rowStyle.setAlignment(HorizontalAlignment.CENTER);

		int rownum = 2;
		for (PendingAndApproveItemResponseDto data : list) {
			Row row = sheet.createRow(rownum++);
			row.setRowStyle(rowStyle);
			createListForPendingAndApprovedItemRequests(data, row);
		}
		workbook.write(out);
		workbook.close();
		out.close();
		return new ByteArrayInputStream(out.toByteArray());
	}

	private void createListForPendingAndApprovedItemRequests(PendingAndApproveItemResponseDto data, Row row) {

		Cell cell = row.createCell(0);
		cell.setCellValue(data.getItemName());
		cell.setCellStyle(row.getRowStyle());

		cell = row.createCell(1);
		cell.setCellValue(data.getItemGroup());
		cell.setCellStyle(row.getRowStyle());

		cell = row.createCell(2);
		cell.setCellValue(data.getRequiredQuantity());
		cell.setCellStyle(row.getRowStyle());

		cell = row.createCell(3);
		cell.setCellValue(data.getRequiredDate());
		cell.setCellStyle(row.getRowStyle());

		cell = row.createCell(4);
		if (data.getRequestDate() == null) {
			cell.setCellValue(" ");
		} else {
			cell.setCellValue(data.getRequestDate());
		}
		cell.setCellStyle(row.getRowStyle());

		cell = row.createCell(5);
		cell.setCellValue(data.getRemarks());
		cell.setCellStyle(row.getRowStyle());

		cell = row.createCell(6);
		cell.setCellValue(data.getAction());
		cell.setCellStyle(row.getRowStyle());

		cell = row.createCell(7);
		cell.setCellValue(data.getRemarkFormStorePR());
		cell.setCellStyle(row.getRowStyle());

		cell = row.createCell(8);
		cell.setCellValue(data.getApprovedDate());
		cell.setCellStyle(row.getRowStyle());

	}

	@Override
	public List<StockInDto> saveStockInList(List<StockInDto> stockInDto) {
		Instant instant = Instant.now();
		Timestamp time = Timestamp.from(instant);
		for (StockInDto list : stockInDto) {
			StoreStockEntryInEo storeStockEntryInEo = new StoreStockEntryInEo();
			storeStockEntryInEo.setItemId(list.getItemId());
			storeStockEntryInEo.setItemGroupId(list.getItemGroupId());
			storeStockEntryInEo.setItemQuantity(list.getQuantity());
			storeStockEntryInEo.setPurchaseDate(list.getPurchaseDate());
			storeStockEntryInEo.setRemarks(list.getRemarks());
			storeStockEntryInEo.setCreatedOn(time);
			stockInRepo.save(storeStockEntryInEo);
		}
		return stockInDto;

	}

	@Override
	public AddBillDto saveAddBill(AddBillDto dtoBill) {
		StoreSectionBillEO billObject = new StoreSectionBillEO();
		Instant instant = Instant.now();
		Timestamp time = Timestamp.from(instant);
		try {
			billObject.setBillDate(dtoBill.getBillDate());
			billObject.setVendorId(dtoBill.getVendorName());
			billObject.setBillType(dtoBill.getBillType());
			billObject.setFileNumber(dtoBill.getFileNumber());
			billObject.setInvoiceNumber(dtoBill.getInvoiceNumber());
			billObject.setGemId(dtoBill.getGemId());
			billObject.setGstNumber(dtoBill.getGstNumber());
			billObject.setAmount(dtoBill.getAmount());
			billObject.setGstOtherTaxes(dtoBill.getGstOtherTaxes());
			billObject.setTotalAmount(dtoBill.getTotalAmount());
			billObject.setCreatedOn(time);
			addBillRepository.save(billObject);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return dtoBill;
	}

	@Override
	public List<DropdownUUIDDto> getAllVendorName() {
		return addBillRepository.getAllVendorName();
	}

	@Override
	public List<DropdownUUIDDto> getAllBillType() {
		return addBillRepository.getAllBillType();
	}

	@Override
	public Page<NonConsumableItemsStockRegisterResponseDto> getListOfNonConsumableItemsStockRegister(
			PendingAndApprovedItemRequestsDto reqDto) {
		Page<NonConsumableItemsStockRegisterResponseDto> list = null;
		try {
			list = dayToDayOperationRepositoryImpl.getListOfNonConsumableItemsStockRegister(reqDto);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;
	}

	@Override
	public Page<ConsumableItemsStockRegisterResponseDto> getListOfConsumableItemsStockRegister(
			PendingAndApprovedItemRequestsDto reqDto) {
		Page<ConsumableItemsStockRegisterResponseDto> list = null;
		try {
			list = dayToDayOperationRepositoryImpl.getListOfConsumableItemsStockRegister(reqDto);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;
	}

	@Override
	public ByteArrayInputStream generateExcelSheetForNonConsumableItemsStockRegister(int pageNumber, int pageSize,
			String itemName, String itemCategory, Date toDate, Date fromDate) throws IOException {
		PendingAndApprovedItemRequestsDto dto = new PendingAndApprovedItemRequestsDto();
		dto.setFromDate(fromDate);
		dto.setToDate(toDate);
		dto.setItemName(itemName);
		dto.setItemCategory(itemCategory);
		dto.setPageNumber(pageNumber);
		dto.setPageSize(pageSize);
		Page<NonConsumableItemsStockRegisterResponseDto> page = dayToDayOperationRepositoryImpl
				.getListOfNonConsumableItemsStockRegister(dto);
		List<NonConsumableItemsStockRegisterResponseDto> list = page.getContent();
		XSSFWorkbook workbook = new XSSFWorkbook();
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		Font font = workbook.createFont();
		font.setColor(IndexedColors.YELLOW.getIndex());
		style.setFont(font);
		XSSFSheet sheet = workbook.createSheet("Non Consumable Items Stock Register");
		Row titleRow = sheet.createRow(0);
		Cell titleCell = titleRow.createCell(0);
		titleCell.setCellValue("Non Consumable Items Stock Register");
		CellStyle titleCellStyle = workbook.createCellStyle();
		Font titleFont = workbook.createFont();
		titleFont.setBold(true);
		titleFont.setFontName("Arial");
		titleCellStyle.setFont(titleFont);
		titleCellStyle.setBorderBottom(BorderStyle.THIN);
		titleCellStyle.setAlignment(HorizontalAlignment.CENTER);
		titleCell.setCellStyle(titleCellStyle);
		Row header = sheet.createRow(1);
		CellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setBorderBottom(BorderStyle.THIN);
		headerStyle.setAlignment(HorizontalAlignment.CENTER);
		headerStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		XSSFFont font1 = ((XSSFWorkbook) workbook).createFont();
		font.setFontName("Arial");
		font.setFontHeightInPoints((short) 10);
		font1.setBold(true);
		headerStyle.setFont(font1);
		Cell headerCell = header.createCell(0);
		headerCell.setCellValue("Year of Account");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(1);
		headerCell.setCellValue("Balance B/F");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(2);
		headerCell.setCellValue("Date of Purchase");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(3);
		headerCell.setCellValue("Bill in which Charged");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(4);
		headerCell.setCellValue("Name of the FIRM");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(5);
		headerCell.setCellValue("No. of Articles Received");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(6);
		headerCell.setCellValue("Total Price");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(7);
		headerCell.setCellValue("For Whose use");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(8);
		headerCell.setCellValue("Total");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(9);
		headerCell.setCellValue("No. of Articles");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(10);
		headerCell.setCellValue("Why Disposed of");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(11);
		headerCell.setCellValue("How Disposed of");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(12);
		headerCell.setCellValue("Sale Price if Any");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(13);
		headerCell.setCellValue("When Credited to Govt.");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(14);
		headerCell.setCellValue("Date of Original Purchase");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(15);
		headerCell.setCellValue("Original Price");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(16);
		headerCell.setCellValue("Balance carried Over");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(17);
		headerCell.setCellValue("Initials of G.O.");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(18);
		headerCell.setCellValue("Remarks");
		headerCell.setCellStyle(headerStyle);

		CellStyle rowStyle = workbook.createCellStyle();
		rowStyle.setAlignment(HorizontalAlignment.CENTER);

		int rownum = 2;
		for (NonConsumableItemsStockRegisterResponseDto data : list) {
			Row row = sheet.createRow(rownum++);
			row.setRowStyle(rowStyle);
			createListForNonConsumableItemsStockRegister(data, row);
		}
		workbook.write(out);
		workbook.close();
		out.close();
		return new ByteArrayInputStream(out.toByteArray());
	}

	private void createListForNonConsumableItemsStockRegister(NonConsumableItemsStockRegisterResponseDto data,
			Row row) {

		Cell cell = row.createCell(0);
		cell.setCellValue(data.getYearOfAccount());

		cell = row.createCell(1);
		cell.setCellValue(data.getBalance());

		cell = row.createCell(2);
		cell.setCellValue(data.getDateOfPurchase());

		cell = row.createCell(3);
		cell.setCellValue(data.getBillInWhichCharged() + "");

		cell = row.createCell(4);
		cell.setCellValue(data.getNameOfTheFirm() + " ");

		cell = row.createCell(5);
		cell.setCellValue(data.getNoOfArticlesReceived());

		cell = row.createCell(6);
		cell.setCellValue(data.getTotalPrice());

		cell = row.createCell(7);
		cell.setCellValue(data.getForWhoseUse());

		cell = row.createCell(8);
		cell.setCellValue(data.getTotal());

		cell = row.createCell(9);
		cell.setCellValue(data.getNoOfArticles());

		cell = row.createCell(10);
		cell.setCellValue(data.getWhyDisposedOf());

		cell = row.createCell(11);
		cell.setCellValue(data.getHowDisposedOf());

		cell = row.createCell(12);
		cell.setCellValue(data.getSalePriceIfAny());

		cell = row.createCell(13);
		cell.setCellValue(data.getWhenCreditedToGovt());

		cell = row.createCell(14);
		cell.setCellValue(data.getDateOfOriginalPurchase());

		cell = row.createCell(15);
		cell.setCellValue(data.getOriginalPrice());

		cell = row.createCell(16);
		cell.setCellValue(data.getBalanceCarriedOver());

		cell = row.createCell(17);
		cell.setCellValue(data.getRemarks());
	}

	@Override
	public ByteArrayInputStream generateExcelSheetForConsumableItemsStockRegister(int pageNumber, int pageSize,
			String itemName, String itemCategory, Date toDate, Date fromDate) throws IOException {
		PendingAndApprovedItemRequestsDto dto = new PendingAndApprovedItemRequestsDto();
		dto.setFromDate(fromDate);
		dto.setToDate(toDate);
		dto.setItemName(itemName);
		dto.setItemCategory(itemCategory);
		dto.setPageNumber(pageNumber);
		dto.setPageSize(pageSize);
		Page<ConsumableItemsStockRegisterResponseDto> page = dayToDayOperationRepositoryImpl
				.getListOfConsumableItemsStockRegister(dto);
		List<ConsumableItemsStockRegisterResponseDto> list = page.getContent();
		XSSFWorkbook workbook = new XSSFWorkbook();
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		Font font = workbook.createFont();
		font.setColor(IndexedColors.YELLOW.getIndex());
		style.setFont(font);
		XSSFSheet sheet = workbook.createSheet("Consumable Items Stock Register");
		Row titleRow = sheet.createRow(0);
		Cell titleCell = titleRow.createCell(0);
		titleCell.setCellValue("Consumable Items Stock Register");
		CellStyle titleCellStyle = workbook.createCellStyle();
		Font titleFont = workbook.createFont();
		titleFont.setBold(true);
		titleFont.setFontName("Arial");
		titleCellStyle.setFont(titleFont);
		titleCellStyle.setBorderBottom(BorderStyle.THIN);
		titleCellStyle.setAlignment(HorizontalAlignment.CENTER);
		titleCell.setCellStyle(titleCellStyle);
		Row header = sheet.createRow(1);
		CellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setBorderBottom(BorderStyle.THIN);
		headerStyle.setAlignment(HorizontalAlignment.CENTER);
		headerStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		XSSFFont font1 = ((XSSFWorkbook) workbook).createFont();
		font.setFontName("Arial");
		font.setFontHeightInPoints((short) 10);
		font1.setBold(true);
		headerStyle.setFont(font1);
		Cell headerCell = header.createCell(0);
		headerCell.setCellValue("Date");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(1);
		headerCell.setCellValue("Particulars");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(2);
		headerCell.setCellValue("Bill No");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(3);
		headerCell.setCellValue("Receipt");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(4);
		headerCell.setCellValue("Issue");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(5);
		headerCell.setCellValue("Balance");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(6);
		headerCell.setCellValue("Remarks");
		headerCell.setCellStyle(headerStyle);

		CellStyle rowStyle = workbook.createCellStyle();
		rowStyle.setAlignment(HorizontalAlignment.CENTER);

		int rownum = 2;
		for (ConsumableItemsStockRegisterResponseDto data : list) {
			Row row = sheet.createRow(rownum++);
			row.setRowStyle(rowStyle);
			createListForConsumableItemsStockRegister(data, row);
		}
		workbook.write(out);
		workbook.close();
		out.close();
		return new ByteArrayInputStream(out.toByteArray());
	}

	private void createListForConsumableItemsStockRegister(ConsumableItemsStockRegisterResponseDto data, Row row) {

		Cell cell = row.createCell(0);
		cell.setCellValue(data.getDate());

		cell = row.createCell(1);
		cell.setCellValue(data.getParticulars());

		cell = row.createCell(2);
		cell.setCellValue(data.getBillNo());

		cell = row.createCell(3);
		cell.setCellValue("hi");

		cell = row.createCell(4);
		cell.setCellValue(data.getIssue() + " ");

		cell = row.createCell(5);
		cell.setCellValue(data.getBalance());

		cell = row.createCell(6);
		cell.setCellValue(data.getRemarks());
	}

	@Override
	public Page<ListOfBillsFromStoreResponseDto> getListOfBillsFromStore(ListOfBillFromStoreRequestDto dto) {
		Page<ListOfBillsFromStoreResponseDto> list = null;
		try {
			list = dayToDayOperationRepositoryImpl.getListOfBillsFromStore(dto);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;
	}

	@Override
	public void updateListOfBillsFromStore(List<UpdateListOfBillsFromStoreRequestDto> requestDto) {
		Instant instant = Instant.now();
		Timestamp updatedTime = Timestamp.from(instant);
		for (UpdateListOfBillsFromStoreRequestDto list : requestDto) {
			StoreSectionBillEO storeSectionBillEO = addBillRepository.findByBillId(list.getBillId());
			System.out.println(storeSectionBillEO);
			storeSectionBillEO.setAction(list.getAction());
			storeSectionBillEO.setPayment_date(list.getPaymentDate());
			storeSectionBillEO.setUtrNo(list.getUtrNo());
			storeSectionBillEO.setUpdatedOn(updatedTime);
			addBillRepository.save(storeSectionBillEO);
		}
	}

	@Override
	public DayToDayHeadMasterEO saveHeadMasterData(DayToDayHeadMasterRequestDto reqDTO) {
		DayToDayHeadMasterEO dtdhm = new DayToDayHeadMasterEO();
		Instant instant = Instant.now();
		Timestamp time = Timestamp.from(instant);
		try {
			dtdhm.setHeadName(reqDTO.getHeadName());
			dtdhm.setHeadCode(reqDTO.getHeadCode());
			dtdhm.setStatus(reqDTO.isStatus());
			dtdhm.setCreatedBy(jwtUtils.getUserFromJwt().getId());
			dtdhm.setCreatedOn(time);
			dayToDayHeadMasterRepository.save(dtdhm);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return dtdhm;
	}

	@Override
	public GrantMasterEO saveGrantMaster(GrantMasterDto req) {
		GrantMasterEO obj = new GrantMasterEO();
		Instant instant = Instant.now();
		Timestamp time = Timestamp.from(instant);
		try {
			obj.setGrantCode(req.getGrantCode());
			obj.setGrantName(req.getGrantName());
			obj.setActive(req.isActive());
			obj.setCreatedOn(time);
			obj.setCreatedBy(jwtUtils.getUserFromJwt().getId());
			masterRepository.save(obj);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return obj;
	}

	@Override
	public void saveEmployeeTransferData(DayToDayEmployeeTransferRequestDto reqDTO) {
		if (reqDTO.getTransferId() == null || reqDTO.getTransferId().equals("")) {

			DayToDayEmployeeTransferEO dtet = new DayToDayEmployeeTransferEO();
			Instant instant = Instant.now();
			Timestamp time = Timestamp.from(instant);
			// DayToDayEmployeeTransferEO dtet = new DayToDayEmployeeTransferEO();
			dtet.setRniEmployee(reqDTO.getIsRniEmployee());
			dtet.setIsNormalTransfer(reqDTO.getIsNormalTransfer());
			dtet.setEmployeeId(reqDTO.getEmployeeId());
			dtet.setEmployeenNewDesignation(reqDTO.getEmployeeNewDesignation());
			dtet.setEmployeeSection(reqDTO.getEmployeeSection());
			dtet.setUploadTransferId(reqDTO.getUploadTransferId());
			dtet.setRelevingRemarks(reqDTO.getRelevingRemarks());
			dtet.setTransferMinistryCode(reqDTO.getTransferMinistryCode());
			dtet.setDateOfreleaving(reqDTO.getDateOfreleaving());
			dtet.setTransferToSection(reqDTO.getTransferToSection());
			dtet.setEmployeeOldDesignation(reqDTO.getEmployeeOldDesignation());
			dtet.setTransferDate(reqDTO.getTransferDate());
			dtet.setCreatedBy(jwtUtils.getUserFromJwt().getId());
			dtet.setCreatedOn(time);
			dayTodayEmployeeTransferRepository.save(dtet);
		} else {
			DayToDayEmployeeTransferEO dtet = dayTodayEmployeeTransferRepository
					.findByTransferId(reqDTO.getTransferId());
			Instant instant = Instant.now();
			Timestamp time = Timestamp.from(instant);
			// DayToDayEmployeeTransferEO dtet = new DayToDayEmployeeTransferEO();
			// dtet.setRniEmployee(reqDTO.getIsRniEmployee());
			// dtet.setIsNormalTransfer(reqDTO.getIsNormalTransfer());
			// dtet.setEmployeeId(reqDTO.getEmployeeId());
			dtet.setEmployeenNewDesignation(reqDTO.getEmployeeNewDesignation());
			dtet.setEmployeeSection(reqDTO.getEmployeeSection());
			dtet.setUploadTransferId(reqDTO.getUploadTransferId());
			dtet.setRelevingRemarks(reqDTO.getRelevingRemarks());
			dtet.setTransferMinistryCode(reqDTO.getTransferMinistryCode());
			dtet.setDateOfreleaving(reqDTO.getDateOfreleaving());
			// dtet.setTransferToSection(reqDTO.getTransferToSection());
			// dtet.setEmployeeOldDesignation(reqDTO.getEmployeeOldDesignation());
			dtet.setTransferDate(reqDTO.getTransferDate());
			dtet.setCreatedBy(jwtUtils.getUserFromJwt().getId());
			dtet.setCreatedOn(time);
			dayTodayEmployeeTransferRepository.save(dtet);
		}

	}

	@Override
	public Page<DayToDaySearchImprestAccResDto> getListOfImprestAccRequestsFromEmployees(
			DayToDaySearchImprestAccReqDto reqDto) {
		Page<DayToDaySearchImprestAccResDto> list = null;
		try {
			list = dayToDayOperationRepositoryImpl.getListOfImprestAccRequestsFromEmployees(reqDto);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;
	}

	@Override
	public Page<DayToDaySearchImprestAccResDto> getListOfImprestAccRaisedAndCompleted(
			DayToDaySearchImprestAccCompletedAndRaisedReqDto reqDto) {
		Page<DayToDaySearchImprestAccResDto> list = null;
		try {
			list = dayToDayOperationRepositoryImpl.getListOfImprestAccRaisedCompleted(reqDto);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;
	}

	@Override
	public List<DayToDayHeadMasterResponseDto> headMasterList() {
		List<DayToDayHeadMasterResponseDto> list = null;
		try {
			list = dayToDayHeadMasterRepository.headMasterList();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;

	}

	@Override
	public void updateHeadMasterList(UpdateHeadMasterRequestDto reqDTO) {
		DayToDayHeadMasterEO daytoday = dayToDayHeadMasterRepository.findById(reqDTO.getHeadId()).get();
		daytoday.setStatus(reqDTO.isStatus());
		dayToDayHeadMasterRepository.save(daytoday);
	}

	@Override

	public void saveEmployeeManagement(EmployeeBasicInfoRequestDto RequestDto) {
		try {
			EmployeeManagementEO empObj = new EmployeeManagementEO();

			if (RequestDto.getEmployeeId() == null || RequestDto.getEmployeeId().equals("")) {

				// empObj.setEmployeeId(UUID.randomUUID());
				empObj.setEmployeeId(jwtUtils.getUserFromJwt().getId());
				empObj.setEmployeeName(RequestDto.getEmployeeName());
				empObj.setEmployeeFileNumber(RequestDto.getEmployeeFileNumber());
				StringBuilder employeeCodeFormate = new StringBuilder();
				employeeCodeFormate.append("RNI/000");
				String priviousemployeeCode = employeeManagementRepository.toCheckEmployeeCode();
				if (priviousemployeeCode != null) {
					String subCodePart = priviousemployeeCode.substring(priviousemployeeCode.indexOf("RNI/000") + 7);
					Integer k = Integer.valueOf(subCodePart);
					Integer j = k + 1;
					employeeCodeFormate.append(j);
					empObj.setEmployeeCode(employeeCodeFormate.toString());
				} else {
					empObj.setEmployeeCode(employeeCodeFormate + "0001");
				}

				empObj.setCastCategoryCode(RequestDto.getCastCategoryCode());
				empObj.setDateOfBirth(RequestDto.getDateOfBirth());
				empObj.setGenderCode(RequestDto.getGenderCode());
				empObj.setNatinality(RequestDto.getCurrentDistrictId());
				empObj.setIsPwd(RequestDto.getIsPwd());
				empObj.setBloodGroupCode(RequestDto.getBloodGroupCode());
				empObj.setMobileNumber(RequestDto.getMobileNumber());
				empObj.setOfficialEmailId(RequestDto.getOfficialEmailId());
				empObj.setEmailId(RequestDto.getEmailId());
				empObj.setPanNumber(RequestDto.getPanNumber());
				empObj.setCurrentAddress(RequestDto.getCurrentAddress());
				empObj.setCurrentStateId(RequestDto.getCurrentStateId());
				empObj.setCurrentDistrictId(RequestDto.getCurrentDistrictId());
				empObj.setCurrentCity(RequestDto.getCurrentCity());
				empObj.setCurrentPinCode(RequestDto.getCurrentPinCode());
				empObj.setPermanentAddress(RequestDto.getPermanentAddress());
				empObj.setPermanentStateId(RequestDto.getCurrentStateId());
				empObj.setPermanentDistrictId(RequestDto.getCurrentDistrictId());
				empObj.setPermanentCity(RequestDto.getPermanentCity());
				empObj.setPermanentPinCode(RequestDto.getPermanentPinCode());
				empObj.setActive(RequestDto.isActive());
				empObj.setCreatedOn(jwtUtils.getTimeFromJwt());
				employeeManagementRepository.save(empObj);

			} else {

				EmployeeManagementEO employeeManagementEO = employeeManagementRepository
						.findByEmployeeId(RequestDto.getEmployeeId());

				employeeManagementEO.setEmployeeId(RequestDto.getEmployeeId());
				employeeManagementEO.setEmployeeName(RequestDto.getEmployeeName());
				employeeManagementEO.setEmployeeFileNumber(RequestDto.getEmployeeFileNumber());
				employeeManagementEO.setEmployeeCode(RequestDto.getEmployeeCode());
				employeeManagementEO.setCastCategoryCode(RequestDto.getCastCategoryCode());
				employeeManagementEO.setDateOfBirth(RequestDto.getDateOfBirth());
				employeeManagementEO.setGenderCode(RequestDto.getGenderCode());
				employeeManagementEO.setNatinality(RequestDto.getCurrentDistrictId());
				employeeManagementEO.setIsPwd(RequestDto.getIsPwd());
				employeeManagementEO.setBloodGroupCode(RequestDto.getBloodGroupCode());
				employeeManagementEO.setMobileNumber(RequestDto.getMobileNumber());
				employeeManagementEO.setOfficialEmailId(RequestDto.getOfficialEmailId());
				employeeManagementEO.setEmailId(RequestDto.getEmailId());
				employeeManagementEO.setPanNumber(RequestDto.getPanNumber());
				employeeManagementEO.setCurrentAddress(RequestDto.getCurrentAddress());
				employeeManagementEO.setCurrentStateId(RequestDto.getCurrentStateId());
				employeeManagementEO.setCurrentDistrictId(RequestDto.getCurrentDistrictId());
				employeeManagementEO.setCurrentCity(RequestDto.getCurrentCity());
				employeeManagementEO.setCurrentPinCode(RequestDto.getCurrentPinCode());
				employeeManagementEO.setPermanentAddress(RequestDto.getPermanentAddress());
				employeeManagementEO.setPermanentStateId(RequestDto.getPermanentStateId());
				employeeManagementEO.setPermanentDistrictId(RequestDto.getCurrentDistrictId());
				employeeManagementEO.setPermanentCity(RequestDto.getPermanentCity());
				employeeManagementEO.setPermanentPinCode(RequestDto.getPermanentPinCode());
				employeeManagementEO.setActive(RequestDto.isActive());
				empObj.setCreatedOn(jwtUtils.getTimeFromJwt());
				employeeManagementRepository.save(employeeManagementEO);
			}
		} catch (Exception e) {

		}

	}

	@Override
	public List<GrantMasterResponseDto> grantMasterList() {
		List<GrantMasterResponseDto> grantMasterlist = null;
		try {
			grantMasterlist = garntMasterListRepository.grantMasterList();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return grantMasterlist;
	}

	@Override
	public void updateGrantMasterList(GrantMasterDto reqDTO) {
		GrantMasterEO update = garntMasterListRepository.findById(reqDTO.getGrantId()).get();
		update.setGrantCode(reqDTO.getGrantCode());
		update.setGrantName(reqDTO.getGrantName());
		update.setActive(reqDTO.isActive());
		garntMasterListRepository.save(update);
	}

	@Override
	public GrantMasterEO findGrrantMasterList(GrantMasterResponseDto dto) {
		GrantMasterEO list = garntMasterListRepository.findById(dto.getGrantId()).get();
		return list;
	}

	@Override
	public DayToDayHeadMasterResponseDto getHeadMasterDetailById(HeadMasterDetailByIdRequesteDto reqDto) {
		return dayToDayHeadMasterRepository.getHeadMasterDetailById(reqDto.getHeadId());
	}

	@Override
	public void saveExpenditureDetail(SaveExpenditureRequestDto reqDto) {
		Instant instant = Instant.now();
		Timestamp createTime = Timestamp.from(instant);
		ExpenditureEo expenditureEo = new ExpenditureEo();
		expenditureEo.setMonthName(reqDto.getMonthName());
		expenditureEo.setYearId(reqDto.getYearId());
		expenditureEo.setSalaryBill(reqDto.getSalaryBill());
		expenditureEo.setOtheEmployee(reqDto.getOtherEmployee());
		expenditureEo.setDepttExpe(reqDto.getDepttExpe());
		expenditureEo.setCreatredOn(createTime);
		expenditureRepository.save(expenditureEo);
	}

	@Override
	public Page<List<ExpenditureListResponseDto>> getListOfExpenditureDetail(PaginationRequestDto reqDto) {

		return expenditureRepository
				.getListOfExpenditureDetail((PageRequest.of(reqDto.getPageIndex(), reqDto.getPageSize())));
	}

	@Override
	public void updateExpenditureDetail(UpdateExpenditureRequestDto reqDto) {
		ExpenditureEo expenditureEo = expenditureRepository.findById(reqDto.getExpenditureId()).get();
		expenditureEo.setDepttExpe(reqDto.getDepttExpe());
		expenditureEo.setMonthName(reqDto.getMonthName());
		expenditureEo.setYearId(reqDto.getYearId());
		expenditureEo.setOtheEmployee(reqDto.getOtherEmployee());
		expenditureEo.setSalaryBill(reqDto.getSalaryBill());
		expenditureRepository.save(expenditureEo);
	}

	@Override
	public ExpenditureResponseDto getExpenditureDetailById(ExpenditureRequestDto reqDto) {
		return expenditureRepository.getExpenditureDetailById(reqDto.getExpenditureId());
	}

	@Override
	public void updateImprestACRequestsFromEmployees(List<UpdateImprestACRequestsFromEmployeeRequestDto> reqDto) {
		Instant instant = Instant.now();
		Timestamp timeStamp = Timestamp.from(instant);
		for (UpdateImprestACRequestsFromEmployeeRequestDto list : reqDto) {
			ImprestRequestEO imprestRequestEO = imprestRepo.findById(list.getImprestId()).get();
			imprestRequestEO.setPaymentDate(list.getPaymentDate());
			imprestRequestEO.setReferenceDocument(list.getReferenceDocument());
			imprestRequestEO.setReferenceDocumentName(list.getReferenceDocumentName());
			imprestRequestEO.setStatus(list.getStatus());
			imprestRequestEO.setModifiedOn(timeStamp);
			imprestRepo.save(imprestRequestEO);
		}
	}

	public Page<List<EmployeeMasterResponseDto>> getEmployeeMasterList(PaginationRequestDto reqDto) {
		Page<List<EmployeeMasterResponseDto>> reqList = null;
		/*
		 * try { reqList = employeeManagementRepository
		 * .getEmployeeMasterList(PageRequest.of(reqDto.getPageIndex(),
		 * reqDto.getPageSize())); } catch (Exception e) { e.printStackTrace();
		 * 
		 * }
		 */
		return reqList;
	}

	@Override
	public EmployeeManagementResponseDto findEmployeeManagementDetailsByEmployeeId(
			EmployeeManagementRequestDto empData) {
		EmployeeManagementResponseDto empDetails = employeeManagementRepository
				.findEmployeeManagementDetailsByEmployeeId(empData.getEmployeeId());
		return empDetails;
	}

	@Override
	public void updateListOfEmployeeMaster(EmployeeManagementDto requestDto) {
		/*
		 * EmployeeManagementEO rsponseData =
		 * employeeManagementRepository.findByEmployeeId(requestDto.getEmployeeId());
		 * rsponseData.setEmployeeCode(requestDto.getEmployeeCode());
		 * rsponseData.setEmployeeFileNumber(requestDto.getEmployeeFileNumber());
		 * rsponseData.setEmployeeName(requestDto.getEmployeeName());
		 * rsponseData.setOfficialEmailId(requestDto.getOfficialEmailId());
		 * rsponseData.setEmailId(requestDto.getEmailId());
		 * rsponseData.setMobileNumber(requestDto.getMobileNumber());
		 * rsponseData.setPanNumber(requestDto.getPanNumber());
		 * rsponseData.setAddharNumber(requestDto.getAddharNumber());
		 * rsponseData.setCurrentAddress(requestDto.getCurrentAddress());
		 * rsponseData.setIsPermanentAddress(requestDto.getIsPermanentAddress());
		 * rsponseData.setPermanentAddress(requestDto.getPermanentAddress());
		 * rsponseData.setDesignationId(requestDto.getDesignationId());
		 * rsponseData.setEmployeeBelongingId(requestDto.getEmployeeBelongingId());
		 * rsponseData.setGroupId(requestDto.getGroupId());
		 * rsponseData.setServiceCadre(requestDto.getServiceCadre());
		 * rsponseData.setGenderId(requestDto.getGenderId());
		 * rsponseData.setDateOfJoiningInRni(requestDto.getDateOfJoiningInRni());
		 * rsponseData.setDateOfJoiningInGovJob(requestDto.getDateOfJoiningInGovJob());
		 * rsponseData.setDateOfRetirement(requestDto.getDateOfRetirement());
		 * rsponseData.setDateOfTransfer(requestDto.getDateOfTransfer());
		 * rsponseData.setEmployeeCategoryId(requestDto.getEmployeeCategoryId());
		 * rsponseData.setIsPwd(requestDto.getIsPwd());
		 * rsponseData.setIsTransferCase(requestDto.getIsTransferCase());
		 * rsponseData.setNormalTransferOnPromotion(requestDto.
		 * getNormalTransferOnPromotion());
		 * rsponseData.setTransferFrom(requestDto.getTransferFrom());
		 * rsponseData.setTransferDepartment(requestDto.getTransferDepartment());
		 * rsponseData.setTransferDesignation(requestDto.getTransferDesignation());
		 * rsponseData.setRelievingDate(requestDto.getRelievingDate());
		 * rsponseData.setSectionRniId(requestDto.getSectionRniId());
		 * rsponseData.setIsBookReceived(requestDto.getIsBookReceived());
		 * rsponseData.setAparDossierGradingYear(new
		 * Gson().toJson(requestDto.getAparDossierGradingYear()));
		 * employeeManagementRepository.save(rsponseData);
		 */
	}

	@Override
	public void saveGrantReceived(GrantReceivedRequestDto reqDto) {
		Instant instant = Instant.now();
		Timestamp createTime = Timestamp.from(instant);
		List<GrantReceivedListRequestDto> data = reqDto.getGrantReceivedListRequestDto();
		for (GrantReceivedListRequestDto list : data) {
			GrantReceivedEo grantReceivedEo = new GrantReceivedEo();
			grantReceivedEo.setHeadId(list.getHeadId());
			grantReceivedEo.setYearId(reqDto.getFinancialYearId());
			grantReceivedEo.setBgSalery(list.getBg());
			grantReceivedEo.setFgSalery(list.getFg());
			grantReceivedEo.setRgSalery(list.getRg());
			grantReceivedEo.setActive(true);
			grantReceivedEo.setCreatedOn(createTime);
			grantReceivedRepository.save(grantReceivedEo);
		}
	}

	public ByteArrayInputStream generateExcelOfEmployeeMaster(EmployeeMasterResponseDto dto) throws IOException {
		List<EmployeeMasterResponseDto> responseData = employeeManagementRepository.generateExcelOfEmployeeMaster(dto);
		XSSFWorkbook workbook = new XSSFWorkbook();
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		Font font = workbook.createFont();
		font.setColor(IndexedColors.YELLOW.getIndex());
		style.setFont(font);
		XSSFSheet sheet = workbook.createSheet("Employee Master List");
		Row titleRow = sheet.createRow(0);
		Cell titleCell = titleRow.createCell(0);
		titleCell.setCellValue("Employee Master List");
		CellStyle titleCellStyle = workbook.createCellStyle();
		Font titleFont = workbook.createFont();
		titleFont.setBold(true);
		titleFont.setFontName("Arial");
		titleCellStyle.setFont(titleFont);
		titleCellStyle.setBorderBottom(BorderStyle.THIN);
		titleCellStyle.setAlignment(HorizontalAlignment.CENTER);
		titleCell.setCellStyle(titleCellStyle);
		Row header = sheet.createRow(1);
		CellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setBorderBottom(BorderStyle.THIN);
		headerStyle.setAlignment(HorizontalAlignment.CENTER);
		headerStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		XSSFFont font1 = ((XSSFWorkbook) workbook).createFont();
		font.setFontName("Arial");
		font.setFontHeightInPoints((short) 10);
		font1.setBold(true);
		headerStyle.setFont(font1);
		Cell headerCell = header.createCell(0);
		headerCell.setCellValue("Employee Name");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(1);
		headerCell.setCellValue("Employee Code");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(2);
		headerCell.setCellValue("Official Email Id");
		headerCell.setCellStyle(headerStyle);

		headerCell = header.createCell(3);
		headerCell.setCellValue("Mobile No");
		headerCell.setCellStyle(headerStyle);

		CellStyle rowStyle = workbook.createCellStyle();
		rowStyle.setAlignment(HorizontalAlignment.CENTER);

		int rownum = 2;
		for (EmployeeMasterResponseDto data : responseData) {
			Row row = sheet.createRow(rownum++);
			row.setRowStyle(rowStyle);
			createExcelForEmployeeMaster(data, row);
		}
		workbook.write(out);
		workbook.close();
		out.close();
		return new ByteArrayInputStream(out.toByteArray());
	}

	private void createExcelForEmployeeMaster(EmployeeMasterResponseDto empData, Row row) {

		Cell cell = row.createCell(0);
		cell.setCellValue(empData.getEmployeeName());

		cell = row.createCell(1);
		cell.setCellValue(empData.getEmployeeCode());

		cell = row.createCell(2);
		cell.setCellValue(empData.getOfficialEmailId());

		cell = row.createCell(3);
		cell.setCellValue(empData.getMobileNumber());

	}

	@Override
	public GetDataOfImprestRequestFormResponseDto getDataOfImprestRequestForm() {
		GetDataOfImprestRequestFormResponseDto response = null;
		/*
		 * try { response =
		 * employeeManagementRepository.getDataOfImprestRequestForm(jwtUtils.
		 * getUserFromJwt().getId()); } catch (Exception e) { e.printStackTrace(); }
		 */
		return response;
	}

	@Override
	public void updatePendingApprovedITEMRequests(List<UpdatePendingApprovedITEMRequestsDto> reqDto) {
		for (UpdatePendingApprovedITEMRequestsDto list : reqDto) {
			StoreSectionItemRequestEO storeSectionItemRequestEO = storeSectionItemRequestRepo
					.findById(list.getRequestId()).get();
			storeSectionItemRequestEO.setStatus(list.getAction());
			storeSectionItemRequestEO.setApprovedDate(list.getApprovedDate());
			storeSectionItemRequestEO.setRemarksFromStorePR(list.getRemarkForStorePr());
			storeSectionItemRequestRepo.save(storeSectionItemRequestEO);
		}
	}

	@Override
	public ItemNameResponseDto getItemNameToItemCatogry(ItemNameRequestDto reqDTO) {
		return itemMasterEntryFormRepo.getItemCategoryByItemName(reqDTO.getItemId());
	}

	@Override
	public VenderGstAndGemNumberResponseDto getVenderGstAndGemNumber(VenderGstAndGemNumberRequestDto reqDto) {
		return addVenderRepository.getVenderGstAndGemNumber(reqDto.getVenderId());
	}

	@Override
	public DayToDayCommonResponseDto getItemCodeItemCatogeryItemGruopBYItemName(ItemNameRequestDto reqDTO) {
		return itemMasterEntryFormRepo.getItemCodeItemCatogeryItemGruopBYItemName(reqDTO.getItemId());
	}

	@Override
	public Page<StockEntryListFormResponseDto> getListOfStockEntryFormFilter(StockEntryListFormRequestDto reqDTO) {
		Page<StockEntryListFormResponseDto> list = null;
		try {
			list = dayToDayOperationRepositoryImpl.getListOfStockEntryFormFilter(reqDTO);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;
	}

	@Override
	public List<StockEntryListFormResponseDto> getListOfStockEntryForm() {
		return itemMasterEntryFormRepo.getListOfStockEntryForm();
	}

	@Override

	public ItemMasterEntryFormDataDto submitItemMasterEntryForm(ItemMasterEntryFormDataDto itemReq) {

		StoreItemMasterEO itemObject = new StoreItemMasterEO();
		Instant instant = Instant.now();
		Timestamp timeStamp = Timestamp.from(instant);
		itemObject.setItemCode(itemReq.getItemCode());
		itemObject.setItemName(itemReq.getItemName());
		itemObject.setItemType(itemReq.getItemType());
		itemObject.setItemGroupId(itemReq.getItemGroup());
		itemObject.setItemCategoryId(itemReq.getItemCategory());
		itemObject.setProcurementId(itemReq.getProcurement());
		itemObject.setUnitOfMeasurementId(itemReq.getUnitOfMeasurementId());
		itemObject.setRemarks(itemReq.getRemarks());
		itemObject.setCreatedOn(timeStamp);
		itemMasterEntryFormRepo.save(itemObject);
		return itemReq;

	}

	@Override
	public Page<ItemMasterEntryFormResponseDTO> getListOfMaterEntryForm(ItemMasterEntryFormListRequestDTO itemReq) {
		Page<ItemMasterEntryFormResponseDTO> list = null;
		try {
			list = dayToDayOperationRepositoryImpl.getListOfMaterEntryForm(itemReq);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;
	}

	@Override
	public GenerateItemCodeResponse generateItemCode(GenerateItemCodeRequestDto dto) {
		GenerateItemCodeResponse response = new GenerateItemCodeResponse();
		StoreItemMasterEO itemMasterObj = new StoreItemMasterEO();
		System.out.println("hi");
		Optional<StoreItemMasterEO> itemName = itemMasterEntryFormRepo.findByItemName(dto.getItemName());
		if (itemName.isPresent()) {
			response.setMessage("Item Already Exists");
			return response;

		} else {
			StringBuilder itemCode = new StringBuilder();
			String firstCharaacter = dto.getItemName().substring(0, 3);
			itemCode.append(firstCharaacter);
			itemCode.append("000");
			String itemCodePrivious = itemMasterEntryFormRepo.checkPriviousItemCode();
			if (itemCodePrivious != null) {
				String str = itemCodePrivious.substring(6);
				Integer i = Integer.valueOf(str);
				Integer j = i + 1;
				itemCode.append(j);
				itemMasterObj.setItemCode(itemCode.toString());

			} else {
				itemMasterObj.setItemCode(firstCharaacter + "000" + "1");
			}
			response.setItemCode(itemCodePrivious);
			response.setMessage("Item Code Generated...");
			itemMasterEntryFormRepo.save(itemMasterObj);

		}
		return response;

	}

	@Override
	public ItemIssueFormDetailResponseDto itemRequestsDetailsIssueFormDetailById(ItemIssueFormDetailRequestDto reqDto) {

		return dayToDayOperationRepository.finditemRequestsDetailsIssueFormDetailById(reqDto.getItemRequestId());

	}

	@Override
	public StockRegisterRequestDto saveStockRegister(StockRegisterRequestDto reqDto) {
		StockRegisterEO stockregister = new StockRegisterEO();
		Authentication authentication = SecurityContextHolder.getContext().getAuthentication();
		UserDetails userDetails = (UserDetails) authentication.getPrincipal();
		UserEO user = userRepository.findByUsername(userDetails.getUsername());
		try {

			stockregister.setItemBalance(reqDto.getItemBalance());
			stockregister.setItemCode(reqDto.getItemCode());
			stockregister.setDepartment(reqDto.getDepartment());
			stockregister.setItemQuantity(reqDto.getItemQuantity());
			stockregister.setTransactionType(reqDto.getTransactionType());
			stockregister.setTrasactionDate(reqDto.getTrasactionDate());
			stockregister.setItemCategoryId(reqDto.getItemCategoryId());
			stockregister.setItemGroupId(reqDto.getItemGroupId());
			dayToDayOperationStockRegisterRepository.save(stockregister);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return reqDto;
	}

	@Override
	public String saveItemGroupMaster(ItemGroupMasterRequestDto reqDTO) {
		ItemGroupMasterEO itemGroup = new ItemGroupMasterEO();
		String status = "";
		/*
		 * Authentication authentication =
		 * SecurityContextHolder.getContext().getAuthentication(); UserDetails
		 * userDetails = (UserDetails) authentication.getPrincipal(); UserEO user =
		 * userRepository.findByUsername(userDetails.getUsername());
		 */
		// UserEO user=userRepository.findById(jwtUtils.getUserFromJwt().getId());
		Instant instant = Instant.now();
		Timestamp createTime = Timestamp.from(instant);
		try {
			// itemGroup.setItemGroupId(reqDTO.getItemGroupId());
			itemGroup.setItemGroupCode(reqDTO.getItemGroupCode());
			itemGroup.setItemGroupName(reqDTO.getItemGroupName());
			itemGroup.setItemCatogeryId(reqDTO.getItemCatogeryId());
			itemGroup.setIsActive(reqDTO.getisActive());
			itemGroup.setCreatedBy(jwtUtils.getUserFromJwt().getId());
			itemGroup.setCreatedOn(createTime);
			itemGroupMasterRepository.save(itemGroup);
			status = "success";
		} catch (Exception e) {
			status = "fail";
			e.printStackTrace();
		}
		return status;

	}

	@Override
	public void updateItemGroupMaster(ItemGroupMasterRequestDto reqDTO) {
		ItemGroupMasterEO update = itemGroupMasterRepository.findById(reqDTO.getItemGroupId()).get();
		update.setItemGroupCode(reqDTO.getItemGroupCode());
		update.setItemGroupName(reqDTO.getItemGroupName());
		update.setItemCatogeryId(reqDTO.getItemCatogeryId());
		update.setIsActive(reqDTO.getIsActive());
		itemGroupMasterRepository.save(update);

	}

	public List<DropdownUUIDDto> searchItemName(SearchItemNameRequestDto reqDto) {
		List<DropdownUUIDDto> list = storeItemRepository.getItemNameSearchList(reqDto.getItemName());
		return list;
	}

	@Override
	public void saveItemTypeMaster(ItemTypeMasterDTO reqDto) {

		if (reqDto.getItemTypeId() == null || reqDto.getItemTypeId().equals("")) {
			ItemTypeMasterEO itemTypeMaster = new ItemTypeMasterEO();
			// itemTypeMaster.setItemTypeId(reqDto.getItemTypeId());
			itemTypeMaster.setItemTypeName(reqDto.getItemTypeName());
			itemTypeMaster.setItemGroupId(reqDto.getItemGroupId());
			itemTypeMaster.setItemCategoryId(reqDto.getItemCategoryId());
			itemTypeMaster.setStatus(reqDto.getStatus());
			itemTypeMaster.setOfficerId(jwtUtils.getUserFromJwt().getId());
			itemTypeMaster.setCreatedOn(jwtUtils.getTimeFromJwt());
			itemTypeMasterRepository.save(itemTypeMaster);
		} else {
			ItemTypeMasterEO getByItemTypeId = itemTypeMasterRepository.findById(reqDto.getItemTypeId()).get();

			getByItemTypeId.setItemTypeId(reqDto.getItemTypeId());
			;
			getByItemTypeId.setItemTypeName(reqDto.getItemTypeName());
			getByItemTypeId.setItemGroupId(reqDto.getItemGroupId());
			getByItemTypeId.setItemCategoryId(reqDto.getItemCategoryId());
			getByItemTypeId.setStatus(reqDto.getStatus());
			itemTypeMasterRepository.save(getByItemTypeId);
		}

	}

	@Override
	public Page<ItemTypeResDto> getListOfItemTypeMaster(ItemTypeMasterDTO reqDto) {
		Page<ItemTypeResDto> list = null;
		try {
			list = dayToDayOperationRepositoryImpl.getListOfItemTypeMaster(reqDto);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;

	}

	@Override
	public void updateItemTypeMaster(ItemTypeMasterDTO reqDto) {
		ItemTypeMasterEO itemTypeMaster = itemTypeMasterRepository.findById(reqDto.getItemTypeId()).get();

		itemTypeMaster.setItemTypeId(reqDto.getItemTypeId());
		;
		itemTypeMaster.setItemTypeName(reqDto.getItemTypeName());
		itemTypeMaster.setItemGroupId(reqDto.getItemGroupId());
		itemTypeMaster.setItemCategoryId(reqDto.getItemCategoryId());
		itemTypeMaster.setStatus(reqDto.getStatus());
		itemTypeMasterRepository.save(itemTypeMaster);
	}

	@Override
	public ItemTypeMasterResponseDTO getItemTypeMasterByID(ItemTypeMasterDTO reqDto) {
		return itemTypeMasterRepository.getItemTypeMasterByID(reqDto.getItemTypeId());

	}

	@Override
	public ItemGroupMasterResponseDto getItemGroupMasterById(ItemGroupMasterRequestDto reqDTO) {

		return itemGroupMasterRepository.getItemGroupMasterById(reqDTO.getItemGroupId());
	}

	@Override
	public Page<ItemGroupMasterResponseDto> getListOfItemGroupMasterList(ItemGroupMasterRequestDto reqDto) {
		Page<ItemGroupMasterResponseDto> list = null;
		try {
			list = dayToDayOperationRepositoryImpl.getListOfItemGroupMasterList(reqDto);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;
	}

	@Override
	public void updateitemRequestDetailsIsuueForm(ItemRequestDetailsIssueForm reqDto) {
		// for (ItemRequestDetailsIssueForm list : reqDto) {
		ITEMRequestFormEO ITEMRequestFormEO = dayToDayOperationRepository.findById(reqDto.getItemRequestId()).get();

		ITEMRequestFormEO.setIssuedQuantity(reqDto.getIssuedQuantity());
		ITEMRequestFormEO.setRemarkForStoreAdmin(reqDto.getRemarksFromStoreAdmin());
		ITEMRequestFormEO.setIssueDate(reqDto.getIssuedDate());
		ITEMRequestFormEO.setActive(reqDto.isAction());

		ITEMRequestFormEO.setSerialNo(reqDto.getSerialNumber());

		dayToDayOperationRepository.save(ITEMRequestFormEO);
	}

	@Override
	public StoreItemRecieveIssueForm getItemRecievedIssueForm(ItemRequestDetailsIssueForm reqDTO) {

		return dayToDayOperationRepository.getItemrecievedIssuFormById(reqDTO.getItemRequestId());
	}

	@Override
	public List<SearchItemNameResponseDto> seachSerialNumber(SearchItemNameRequestDto reqDto) {
		List<SearchItemNameResponseDto> list = dayToDayOperationRepository.searchSerialNumber(reqDto.getSerialNumber());

		return list;
	}

	@Override
	public DayToDayEmployeetransfeResDto getDesignationPostedInSectionByEmployeeName(
			DayToDayEmployeeTransferRequestDto reqDto) {

		return dayTodayEmployeeTransferRepository.getDesignationPostedInSectionByEmployeeName();
	}

	@Override
	public Page<DayToDayEmployeetransfeResDto> getListOfEmployeeTransfer(DayToDayEmployeeTransferRequestDto reqDto) {

		Page<DayToDayEmployeetransfeResDto> list = null;
		try {
			list = dayToDayOperationRepositoryImpl.getEmployeeTransferList(reqDto);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;
	}

	@Override
	public DayToDayEmployeeTransferViewResDto getListEmployeeTransferById(DayToDayEmployeeTransferRequestDto reqDTO) {

		return dayTodayEmployeeTransferRepository.getListOfEmployeeTransferById(reqDTO.getTransferId());

	}

	public void saveEmployeeProfessionalInformation(@Valid EmployeeProfessionalDetailsRequestDto dto) {

		if (dto.getEmployeeId() == null || dto.getEmployeeId().equals("")) {

			EmployeeManagementEO employeeManagementEO = employeeManagementRepository
					.findByEmployeeId(jwtUtils.getUserFromJwt().getId());

			employeeManagementEO.setDesignationId(dto.getDesignationId());
			employeeManagementEO.setDepartmentId(dto.getDepartmentId());
			employeeManagementEO.setSectionCode(dto.getSectionCode());
			employeeManagementEO.setEmployeeBelongingId(dto.isEmployeeBelongingId());
			employeeManagementEO.setGroupCode(dto.getGroupCode());
			employeeManagementEO.setServiceCadre(dto.getServiceCadre());
			employeeManagementEO.setDateOfJoiningInGovJob(dto.getDateOfJoiningInGovJob());
			employeeManagementEO.setDateOfJoiningInRni(dto.getDateOfJoiningInRni());
			employeeManagementEO.setDateOfRetirement(dto.getDateOfRetirement());
			employeeManagementEO.setTransferCase(dto.isTransferCase());
			employeeManagementEO.setNormalTransferOnPromotion(dto.getNormalTransferOnPromotion());
			employeeManagementEO.setTransferFrom(dto.getTransferFrom());
			employeeManagementEO.setDateOfTransfer(dto.getDateOfTransfer());
			employeeManagementEO.setPreviousServiceCadreCode(dto.getPreviousServiceCadreCode());
			employeeManagementEO.setPreviousDesignation(dto.getPreviousDesignation());
			employeeManagementEO.setPriviousGroupCode(dto.getPriviousGroupCode());
			employeeManagementEO.setTransferDepartmentId(dto.getTransferDepartmentId());
			employeeManagementEO.setPostedInSectionCode(dto.getPostedInSectionCode());
			employeeManagementEO.setTransferDesignationId(dto.getTransferDesignationId());

			employeeManagementRepository.save(employeeManagementEO);
		} else {
			EmployeeManagementEO employeeManagementEO = employeeManagementRepository
					.findByEmployeeId(dto.getEmployeeId());

			employeeManagementEO.setEmployeeId(dto.getEmployeeId());
			employeeManagementEO.setDesignationId(dto.getDesignationId());
			employeeManagementEO.setDepartmentId(dto.getDepartmentId());
			employeeManagementEO.setSectionCode(dto.getSectionCode());
			employeeManagementEO.setEmployeeBelongingId(dto.isEmployeeBelongingId());
			employeeManagementEO.setGroupCode(dto.getGroupCode());
			employeeManagementEO.setServiceCadre(dto.getServiceCadre());
			employeeManagementEO.setDateOfJoiningInGovJob(dto.getDateOfJoiningInGovJob());
			employeeManagementEO.setDateOfJoiningInRni(dto.getDateOfJoiningInRni());
			employeeManagementEO.setDateOfRetirement(dto.getDateOfRetirement());
			employeeManagementEO.setTransferCase(dto.isTransferCase());
			employeeManagementEO.setTransferCase(dto.isTransferCase());
			employeeManagementEO.setTransferFrom(dto.getTransferFrom());
			employeeManagementEO.setDateOfTransfer(dto.getDateOfTransfer());
			employeeManagementEO.setPreviousServiceCadreCode(dto.getPreviousServiceCadreCode());
			employeeManagementEO.setPreviousDesignation(dto.getPreviousDesignation());
			employeeManagementEO.setPriviousGroupCode(dto.getPriviousGroupCode());
			employeeManagementEO.setTransferDepartmentId(dto.getTransferDepartmentId());
			employeeManagementEO.setPostedInSectionCode(dto.getPostedInSectionCode());
			employeeManagementEO.setTransferDesignationId(dto.getTransferDesignationId());

			employeeManagementRepository.save(employeeManagementEO);

		}

	}

	@Override
	public void saveEmployeeAPARDetails(@Valid AparDetailsRequestDto requestDto) {
		if (requestDto.getEmployeeId() == null || requestDto.getEmployeeId().equals("")) {

			EmployeeManagementEO employeeManagementEO = employeeManagementRepository
					.findByEmployeeIdWithAparDetailsEagerly(jwtUtils.getUserFromJwt().getId());

			employeeManagementEO.setAprDetailsReceived(requestDto.getIsAprDetailsReceived());
			employeeManagementEO.setServiceBookReceived(requestDto.getIsServiceBookReceived());
			employeeManagementEO.setAprStartDate(requestDto.getAprStartDate());
			employeeManagementEO.setAprEndDate(requestDto.getAprEndDate());
			Set<APARDetailsEO> aparDetailsEOList = employeeManagementEO.getAparDossierGradingYearWise();
			if (aparDetailsEOList == null) {

				aparDetailsEOList = new HashSet<>();
				employeeManagementEO.setAparDossierGradingYearWise(aparDetailsEOList);
			} else {
				// Clear previously added child entries
				aparDetailsEOList.clear();
			}
			for (APARDetailsEO detail : requestDto.getAparDossierGradingYearWise()) {
				APARDetailsEO aparDetails = new APARDetailsEO();
				aparDetails.setEmployeeId(jwtUtils.getUserFromJwt().getId());
				aparDetails.setYearId(detail.getYearId());
				aparDetails.setGrade(detail.getGrade());
				// Establishing the relationship with the employee
				aparDetails.setEmployee(employeeManagementEO);

				aparDetailsRepository.save(aparDetails);
				aparDetailsEOList.add(aparDetails);
			}
			employeeManagementRepository.save(employeeManagementEO);

		} else {
			EmployeeManagementEO employeeManagementEO = employeeManagementRepository
					.findByEmployeeIdWithAparDetailsEagerly(jwtUtils.getUserFromJwt().getId());

			employeeManagementEO.setEmployeeId(requestDto.getEmployeeId());
			employeeManagementEO.setAprDetailsReceived(requestDto.getIsAprDetailsReceived());
			employeeManagementEO.setServiceBookReceived(requestDto.getIsServiceBookReceived());
			employeeManagementEO.setAprStartDate(requestDto.getAprStartDate());
			employeeManagementEO.setAprEndDate(requestDto.getAprEndDate());
			Set<APARDetailsEO> aparDetailsEOList = employeeManagementEO.getAparDossierGradingYearWise();
			if (aparDetailsEOList == null) {

				aparDetailsEOList = new HashSet<>();
				employeeManagementEO.setAparDossierGradingYearWise(aparDetailsEOList);
			} else {
				// Clear previously added child entries
				aparDetailsEOList.clear();
			}
			for (APARDetailsEO detail : requestDto.getAparDossierGradingYearWise()) {
				APARDetailsEO aparDetails = new APARDetailsEO();
				aparDetails.setEmployeeId(jwtUtils.getUserFromJwt().getId());
				aparDetails.setYearId(detail.getYearId());
				aparDetails.setGrade(detail.getGrade());
				// Establishing the relationship with the employee
				aparDetails.setEmployee(employeeManagementEO);

				aparDetailsRepository.save(aparDetails);
				aparDetailsEOList.add(aparDetails);
			}
			employeeManagementRepository.save(employeeManagementEO);
		}
	}

	@Override
	public LocalDate calculateAgeOfRetirement(@Valid EmployeeManagementDto requestDto) {
		int retirementAge = 60;
		LocalDate dateOfBirth = requestDto.getDateOfBirth();
		// LocalDate dateOfJoining = requestDto.getDateOfJoiningInRni();
		/*
		 * Period presentDate = Period.between(dateOfBirth, dateOfJoining);
		 * System.out.println(presentDate); int age = presentDate.getYears();
		 */
		// int remainingYearForRetirement = 60-age ;
		LocalDate retirementDate = dateOfBirth.plusYears(retirementAge);
		return retirementDate;

	}

	@Override
	public EmployeeManagementEO getViewOfEmployeeManagement(@Valid EmployeeManagementDto requestDto) {
		return employeeManagementRepository.findByEmployeeId(requestDto.getEmployeeId());
	}

	@Override
	public Page<EmployeeManagementDto> getEmployeeManagementListPage(@Valid EmployeeManagementDto requestDto) {
		Page<EmployeeManagementDto> list = null;
		try {
			list = dayToDayOperationRepositoryImpl.getEmployeeManagementListPage(requestDto);
		} catch (Exception e) {
			e.printStackTrace();
		}
		return list;

	}

	@Override
	public void UpadteListOfRaisedAndApprovedRequestForSo(PendingAndApprovedItemRequestsDto reqDto) {
		StoreSectionItemRequestEO storeSectionItemRequestEO = storeSectionItemRequestRepo
				.findById(reqDto.getRequestId()).get();
		storeSectionItemRequestEO.setApprovedQuanity(reqDto.getApprovedQuantity());
		storeSectionItemRequestEO.setRemarksFromStoreSO(reqDto.getRemarkFormStoreSO());
		storeSectionItemRequestEO.setStatus(reqDto.getAction());
		storeSectionItemRequestRepo.save(storeSectionItemRequestEO);
	}
}