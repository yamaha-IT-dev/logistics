SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE spGetShipmentDetails


@intShipmentID		NUMERIC

AS

BEGIN

	SELECT * FROM tbl_shipments WHERE shipment_id= @intShipmentID

END
GO

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO