package reference.dto;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;
import reference.file.ExcelDto;

@Getter
@Setter
@AllArgsConstructor
public class TestUser2 implements ExcelDto {
    private TestUser testuser1;
    private TestUser testuser2;
}
